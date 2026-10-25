"""
Автопостинг у Threads: бере випадковий текст із Word-файлу в папці posts/
і випадковий скріншот із папки screenshots/ та публікує їх у Threads.
Без повторів, поки не будуть використані всі тексти (і окремо всі скріншоти),
після чого цикл починається знову.

Запускається через GitHub Actions (див. .github/workflows/autopost.yml).
"""

import hashlib
import json
import os
import random
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from docx import Document

ROOT = Path(__file__).resolve().parent
POSTS_DIR = ROOT / "posts"
IMAGES_DIR = ROOT / "screenshots"
STATE_FILE = ROOT / "state.json"

API = "https://graph.threads.net/v1.0"
TZ = ZoneInfo("Europe/Warsaw")
MORNING_HOUR = 9    # ранкова публікація (за польським часом)
EVENING_HOUR = 19   # вечірня публікація
MAX_LEN = 490       # ліміт Threads — 500 символів, беремо із запасом
SEPARATORS = {"---", "***", "___", "—", "–––"}
IMAGE_EXT = {".jpg", ".jpeg", ".png"}


# ---------- Тексти з Word ----------

def load_posts():
    files = sorted(p for p in POSTS_DIR.glob("*.docx") if not p.name.startswith("~$"))
    if not files:
        sys.exit("❌ У папці posts/ немає Word-файлу (.docx).")

    posts = []
    for f in files:
        paragraphs = [p.text.rstrip() for p in Document(f).paragraphs]
        has_separator = any(t.strip() in SEPARATORS for t in paragraphs)

        current = []
        for t in paragraphs:
            is_break = t.strip() in SEPARATORS if has_separator else t.strip() == ""
            if is_break:
                if current:
                    posts.append("\n".join(current).strip())
                current = []
            elif t.strip() or has_separator:
                current.append(t)
        if current:
            posts.append("\n".join(current).strip())

    posts = [p for p in posts if p]
    print(f"📄 Знайдено постів: {len(posts)}")
    return posts


def load_images():
    images = sorted(p.name for p in IMAGES_DIR.iterdir() if p.suffix.lower() in IMAGE_EXT)
    if not images:
        sys.exit("❌ У папці screenshots/ немає картинок (.jpg / .png).")
    print(f"🖼  Знайдено скріншотів: {len(images)}")
    return images


def post_id(text):
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12]


# ---------- Стан (що вже використано) ----------

def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {"used_posts": [], "used_images": [], "last_slot": "", "history": []}


def save_state(state):
    state["history"] = state["history"][-60:]
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def pick(items, used, key=lambda x: x):
    """Випадковий елемент, який ще не використовувався в цьому циклі."""
    fresh = [i for i in items if key(i) not in used]
    if not fresh:  # усе використано — новий цикл
        used.clear()
        fresh = list(items)
        print("🔄 Усе використано — починаємо новий цикл.")
    choice = random.choice(fresh)
    used.append(key(choice))
    return choice


# ---------- Розбиття довгого тексту ----------

def split_text(text, limit=MAX_LEN):
    parts = []
    while len(text) > limit:
        cut = -1
        for sep in ("\n\n", "\n", ". ", "! ", "? ", " "):
            cut = text.rfind(sep, 0, limit)
            if cut > limit // 3:
                cut += len(sep)
                break
        if cut <= 0:
            cut = limit
        parts.append(text[:cut].strip())
        text = text[cut:].strip()
    if text:
        parts.append(text)
    return parts


# ---------- Threads API ----------

class Threads:
    def __init__(self, token):
        self.token = token
        r = requests.get(f"{API}/me", params={"fields": "id,username", "access_token": token}, timeout=30)
        self._check(r)
        me = r.json()
        self.user_id = me["id"]
        print(f"👤 Акаунт: @{me.get('username')}")

    @staticmethod
    def _check(r):
        if r.status_code != 200:
            sys.exit(f"❌ Помилка Threads API ({r.status_code}): {r.text}")

    def _wait(self, container_id):
        for _ in range(36):  # до 3 хвилин
            time.sleep(5)
            r = requests.get(f"{API}/{container_id}",
                             params={"fields": "status,error_message", "access_token": self.token}, timeout=30)
            self._check(r)
            status = r.json().get("status")
            if status == "FINISHED":
                return
            if status in ("ERROR", "EXPIRED"):
                sys.exit(f"❌ Threads не зміг обробити пост: {r.json()}")
        sys.exit("❌ Threads надто довго обробляє пост.")

    def publish(self, text, image_url=None, reply_to=None):
        data = {"access_token": self.token, "text": text}
        if image_url:
            data.update(media_type="IMAGE", image_url=image_url)
        else:
            data["media_type"] = "TEXT"
        if reply_to:
            data["reply_to_id"] = reply_to

        r = requests.post(f"{API}/{self.user_id}/threads", data=data, timeout=60)
        self._check(r)
        container = r.json()["id"]
        self._wait(container)

        r = requests.post(f"{API}/{self.user_id}/threads_publish",
                          data={"creation_id": container, "access_token": self.token}, timeout=60)
        self._check(r)
        return r.json()["id"]


# ---------- Головна логіка ----------

def current_slot(now):
    if now.hour >= EVENING_HOUR:
        return f"{now:%Y-%m-%d}-evening"
    if now.hour >= MORNING_HOUR:
        return f"{now:%Y-%m-%d}-morning"
    return None


def main():
    force = os.environ.get("FORCE", "false").lower() == "true"
    now = datetime.now(TZ)
    state = load_state()

    slot = current_slot(now)
    if not force:
        if slot is None:
            print(f"⏰ {now:%H:%M} — ще не час публікації.")
            return
        if state["last_slot"] == slot:
            print(f"✅ Пост для {slot} вже опубліковано, пропускаємо.")
            return

    token = os.environ.get("THREADS_TOKEN")
    if not token:
        sys.exit("❌ Не задано секрет THREADS_TOKEN у налаштуваннях GitHub.")

    posts = load_posts()
    images = load_images()

    text = pick(posts, state["used_posts"], key=post_id)
    image = pick(images, state["used_images"])

    repo = os.environ["GITHUB_REPOSITORY"]
    branch = os.environ.get("GITHUB_REF_NAME", "main")
    image_url = f"https://raw.githubusercontent.com/{repo}/{branch}/screenshots/{urllib.parse.quote(image)}"

    print(f"📝 Текст: {text[:80]!r}...")
    print(f"🖼  Скріншот: {image}")

    api = Threads(token)
    parts = split_text(text)
    first_id = api.publish(parts[0], image_url=image_url)
    prev = first_id
    for part in parts[1:]:  # якщо текст довший за 500 символів — продовження у відповідях
        time.sleep(3)
        prev = api.publish(part, reply_to=prev)

    if slot:
        state["last_slot"] = slot
    state["history"].append({"time": now.isoformat(timespec="minutes"), "post": post_id(text),
                             "image": image, "threads_id": first_id})
    save_state(state)
    print(f"🎉 Опубліковано! ID посту: {first_id} (частин: {len(parts)})")


if __name__ == "__main__":
    main()
