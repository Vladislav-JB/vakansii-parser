"""Разовый перенос накопленной базы постов в Supabase.

Запускается НА СВОЁМ КОМПЬЮТЕРЕ (не в облаке), потому что читает локальные файлы:
    python seed_supabase.py "C:\\...\\Парсер Telegram-вакансий\\data\\it_posts.json"
                            ["...\\data\\it_state.json"]
Переменные окружения: SUPABASE_URL и SUPABASE_SERVICE_KEY.

Все посты заливаются с notified=true, чтобы старые вакансии не ушли в бот.
Скрипт можно запускать повторно: уже существующие посты не трогаются.
"""
import hashlib
import json
import os
import sys
import time

import requests

from export_table import norm_words
from label_posts import clean_text
import it_monitor as core  # только ради strip_footer

SB_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SB_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
if not SB_URL or not SB_KEY:
    print("Задай переменные окружения SUPABASE_URL и SUPABASE_SERVICE_KEY.")
    sys.exit(1)


def fingerprint_of_post(p):
    """Тот же отпечаток, что в cloud_monitor.py, чтобы дедуп совпадал."""
    body = clean_text(core.strip_footer(p["text"]))[0]
    return hashlib.sha1(" ".join(sorted(set(norm_words(body)))).encode("utf-8")).hexdigest()


def sb_request(method, path, params=None, body=None, prefer=None):
    headers = {"apikey": SB_KEY, "Authorization": f"Bearer {SB_KEY}",
               "Content-Type": "application/json"}
    if prefer:
        headers["Prefer"] = prefer
    r = requests.request(method, f"{SB_URL}/rest/v1/{path}", headers=headers,
                         params=params, json=body, timeout=120)
    if r.status_code >= 300:
        raise RuntimeError(f"Supabase {method} {path}: HTTP {r.status_code}: {r.text[:300]}")
    return r


def main():
    posts_path = sys.argv[1]
    store = json.load(open(posts_path, encoding="utf-8"))
    rows = [{"post_url": p["url"], "channel": p["channel"], "post_id": p["post_id"],
             "date": p["date"], "text": p["text"], "links": p["links"],
             "buttons": p.get("buttons", []), "fingerprint": fingerprint_of_post(p),
             "notified": True} for p in store.values()]
    print(f"Постов к заливке: {len(rows)}", flush=True)
    t0 = time.time()
    for i in range(0, len(rows), 200):
        sb_request("POST", "vp_posts", prefer="resolution=ignore-duplicates,return=minimal",
                   body=rows[i:i + 200])
        print(f"  залито {min(i + 200, len(rows))} из {len(rows)}", flush=True)
    if len(sys.argv) > 2:
        state = json.load(open(sys.argv[2], encoding="utf-8"))
        srows = [{"channel": ch, "last_seen": ls, "updated_at": __import__("datetime").datetime.now(
            __import__("datetime").timezone.utc).isoformat()} for ch, ls in state.items()]
        for i in range(0, len(srows), 200):
            sb_request("POST", "vp_state", prefer="resolution=merge-duplicates",
                       body=srows[i:i + 200])
        print(f"Состояние каналов: {len(srows)}", flush=True)
    print(f"Готово за {time.time() - t0:.0f} с", flush=True)


if __name__ == "__main__":
    main()
