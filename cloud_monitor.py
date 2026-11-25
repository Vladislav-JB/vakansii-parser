"""Облачная версия парсера вакансий Telegram: GitHub Actions + Supabase.

Отличия от локального it_monitor.py:
  - состояние (последний прочитанный пост каждого канала) и все накопленные посты
    лежат в таблицах Supabase vp_state и vp_posts, а не в локальных JSON-файлах;
  - уведомления уходят напрямую в бот через Bot API (без tg.ps1);
  - таблица xlsx в облаке не собирается: посты копятся в Supabase, выгрузку можно
    сделать позже отдельным шагом.

Запуск (в GitHub Actions, значения из секретов репозитория):
  SUPABASE_URL, SUPABASE_SERVICE_KEY, TG_BOT_TOKEN, TG_CHAT_ID
Логика чтения, фильтров и оценок та же, что у локального парсера (модуль it_monitor.py).
Каждые 10 минут обрабатывается одна из двух групп каналов (по чётности позиции в
channels.txt). Группа выбирается по счётчику в Supabase (как data/it_part.txt у
локальной задачи): 1 -> 2 -> 1 -> ..., чтобы задержки и отмены прогонов не сдвигали
очередь.
"""
import hashlib
import html as html_lib
import json
import os
import random
import re
import sys
import time
from collections import Counter
from datetime import datetime, timezone

import requests

from export_table import jaccard, norm_words, shingles
from fetch_posts import UA, parse_page
from label_posts import clean_text
import it_monitor as core  # темы, оценки, фильтры, reply_link и прочая логика парсера

SB_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SB_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
TG_TOKEN = os.environ.get("TG_BOT_TOKEN", "")
TG_CHAT = os.environ.get("TG_CHAT_ID", "")
if not SB_URL or not SB_KEY:
    print("Нет SUPABASE_URL или SUPABASE_SERVICE_KEY — запуск бессмыслен.")
    sys.exit(1)

BOILER_WINDOW = 200   # сколько последних постов канала смотрим при поиске рекламных ссылок
BATCH = 200           # размер пачки при записи в Supabase
PART_KEY = "__part__"  # зарезервированный канал в vp_state для счётчика групп


# ---------- Supabase (REST/PostgREST) ----------

def sb_request(method, path, params=None, body=None, prefer=None):
    """Один запрос к REST API Supabase. Ответ не 2xx — исключение с текстом ошибки."""
    headers = {"apikey": SB_KEY, "Authorization": f"Bearer {SB_KEY}",
               "Content-Type": "application/json"}
    if prefer:
        headers["Prefer"] = prefer
    r = requests.request(method, f"{SB_URL}/rest/v1/{path}", headers=headers,
                         params=params, json=body, timeout=60)
    if r.status_code >= 300:
        raise RuntimeError(f"Supabase {method} {path}: HTTP {r.status_code}: {r.text[:300]}")
    return r


def load_state():
    rows = sb_request("GET", "vp_state", params={"select": "channel,last_seen"}).json()
    return {row["channel"]: row["last_seen"] for row in rows if row["channel"] != PART_KEY}


def save_state(ch, last_seen):
    sb_request("POST", "vp_state", prefer="resolution=merge-duplicates",
               body={"channel": ch, "last_seen": last_seen,
                     "updated_at": datetime.now(timezone.utc).isoformat()})


def load_part():
    """Текущая группа (1..4) из счётчика в vp_state. Нет записи / мусор — начинаем с 1."""
    rows = sb_request("GET", "vp_state", params={
        "select": "last_seen", "channel": f"eq.{PART_KEY}", "limit": 1}).json()
    try:
        v = int(rows[0]["last_seen"])
        return v if 1 <= v <= 4 else 1
    except (IndexError, TypeError, ValueError):
        return 1


def save_part(part):
    sb_request("POST", "vp_state", prefer="resolution=merge-duplicates",
               body={"channel": PART_KEY, "last_seen": part,
                     "updated_at": datetime.now(timezone.utc).isoformat()})


def insert_posts(rows):
    """Дописывает посты в vp_posts. Уже существующие не трогает (не перезаписывает notified)."""
    for i in range(0, len(rows), BATCH):
        sb_request("POST", "vp_posts", prefer="resolution=ignore-duplicates,return=minimal",
                   body=rows[i:i + BATCH])


def load_pending():
    """Посты прошлых прогонов, которые ещё не отправлены в бот (например, сеть упала)."""
    return sb_request("GET", "vp_posts", params={
        "select": "post_url,channel,post_id,date,text,links,buttons",
        "notified": "eq.false", "order": "post_id.asc", "limit": 200}).json()


def mark_notified(urls):
    for i in range(0, len(urls), BATCH):
        sb_request("PATCH", "vp_posts", params={"post_url": f"in.({','.join(urls[i:i + BATCH])})"},
                   body={"notified": True})


# ---------- Дедупликация ----------

def fingerprint_of_post(p):
    """Отпечаток текста поста: хэш набора слов без ссылок, телефонов и ников.
    Точный, но дешёвый (вместо попарного jaccard по всей базе)."""
    body = clean_text(core.strip_footer(p["text"]))[0]
    return hashlib.sha1(" ".join(sorted(set(norm_words(body)))).encode("utf-8")).hexdigest()


def known_fingerprints(fps, exclude_urls):
    """Какие из отпечатков уже встречались у СТАРЫХ постов (не из этого прогона)."""
    found = set()
    fps = sorted(fps)
    exclude = sorted(exclude_urls)
    for i in range(0, len(fps), 100):
        params = {"select": "fingerprint", "fingerprint": f"in.({','.join(fps[i:i + 100])})"}
        if exclude:
            params["post_url"] = f"not.in.({','.join(exclude)})"
        for row in sb_request("GET", "vp_posts", params=params).json():
            found.add(row["fingerprint"])
    return found


# ---------- Фильтры и оценки (логика как в it_monitor.build_records) ----------

def build_record(p, boiler, stats):
    """Из одного поста делает запись, если он подходит. Иначе None."""
    if core.is_digest(p["text"]):
        stats["подборок"] += 1
        return None
    body = core.strip_footer(p["text"])
    topics = core.find_topics(body)
    if not topics:
        stats["без тем"] += 1
        return None
    text, had_phone, had_email = clean_text(body)
    title = core.make_title(text)
    reason = core.stop_reason(text)
    if p["date"] and (datetime.now(timezone.utc) - datetime.fromisoformat(p["date"])).days > core.MAX_AGE_DAYS:
        reason = f"старый пост (больше {core.MAX_AGE_DAYS} дней)"
    for name, _href in p.get("buttons", []):
        if core.CLOSED_BUTTON_RE.search(name):
            reason = f"закрыто: кнопка «{name}» (на момент чтения)"
    if reason:
        stats["стоп-слова"] += 1
        return None
    score, score_notes = core.score_post(title, text)
    best, _ = core.reply_link(p, boiler)
    if core.KWORK_ORDER_RE.search(best):
        old_kwork = False
        if p["date"]:
            age_h = (datetime.now(timezone.utc) - datetime.fromisoformat(p["date"])).total_seconds() / 3600
            old_kwork = not core.INCLUDE_KWORK or age_h > core.KWORK_MAX_AGE_HOURS
        else:
            old_kwork = not core.INCLUDE_KWORK
        if old_kwork:
            stats["заказы Kwork"] += 1
            return None
    if not best:
        best = "в посте нет ссылки (см. пост; там телефон/e-mail)" if (had_phone or had_email) \
            else "в посте нет ссылки (см. пост)"
    return {
        "date": p["date"], "channel": p["channel"], "url": p["url"], "reply": best, "title": title,
        "topics": ", ".join(topics), "text": text, "score": score, "score_notes": score_notes,
        "tags": core.tags_of(text),
    }


def boilerplate_links(ch):
    """Ссылки-реклама канала: встречаются в >= BOILERPLATE_MIN из последних постов."""
    rows = sb_request("GET", "vp_posts", params={
        "select": "links", "channel": f"eq.{ch}", "order": "post_id.desc",
        "limit": BOILER_WINDOW}).json()
    count = Counter()
    for row in rows:
        for l in {html_lib.unescape(x) for x in (row["links"] or [])}:
            count[l] += 1
    return {l for l, n in count.items() if n >= core.BOILERPLATE_MIN}


# ---------- Уведомления (Bot API напрямую) ----------

def tg_send(text):
    if not TG_TOKEN or not TG_CHAT:
        print("Нет TG_BOT_TOKEN/TG_CHAT_ID — сообщение не отправлено:\n" + text[:400])
        return True  # тестовый режим: не держим посты в pending
    r = requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage",
                      json={"chat_id": TG_CHAT, "text": text, "disable_web_page_preview": True},
                      timeout=30)
    if r.status_code != 200:
        print("Telegram ответил", r.status_code, ":", r.text[:200])
    return r.status_code == 200


def notify_new(records):
    """Шлёт лучшие записи в бот (порциями до ~3500 символов, как локальный парсер).
    Возвращает ссылки записей, которые УДАЛОСЬ отправить (остальные останутся в pending)."""
    records = sorted(records, key=lambda r: (r["score"], r["date"]), reverse=True)
    items = []
    for r in records[:core.NOTIFY_MAX_ITEMS]:
        desc = re.sub(r"\s+", " ", r["text"]).strip()[:200]
        tags = ", ".join(r["tags"])
        items.append((r["url"], f"{r['title']}\nОценка {r['score']} · {r['channel']} · {r['date'][:10]}"
                     + (f"\nТеги: {tags}" if tags else "")
                     + f"\nОтклик: {r['reply']}\nПост: {r['url']}\n{desc}…"))
    sent, header, chunk, chunk_urls = [], f"🔎 Новые вакансии и заказы (оценка {core.NOTIFY_MIN_SCORE}+): {len(records)}\n\n", "", []
    for i, (url, it) in enumerate(items, 1):
        if len(chunk) + len(it) > 3500 and chunk:
            if tg_send(header + chunk):
                sent += chunk_urls
            header, chunk, chunk_urls = "(продолжение)\n\n", "", []
        chunk += f"{i}. {it}\n\n"
        chunk_urls.append(url)
    if chunk:
        if len(records) > core.NOTIFY_MAX_ITEMS:
            chunk += f"…и ещё {len(records) - core.NOTIFY_MAX_ITEMS}, полный список копится в Supabase"
        if tg_send(header + chunk):
            sent += chunk_urls
    return sent


# ---------- Основной обход ----------

def fetch_channel(session, ch, last_seen, made):
    """Читает новые посты одного канала. Возвращает (свежие посты, число запросов)."""
    collected, before = {}, None
    for page in range(core.MAX_PAGES):
        if made[0]:
            time.sleep(random.uniform(*core.PAUSE))
        url = f"https://t.me/s/{ch}" + (f"?before={before}" if before else "")
        resp = core.get_with_retry(session, url)
        made[0] += 1
        print(f"{ch} стр.{page + 1}: HTTP {resp.status_code}", flush=True)
        if resp.status_code != 200:
            raise RuntimeError(f"Telegram ответил кодом {resp.status_code} на канал {ch}")
        posts = parse_page(resp.text, ch)
        if not posts:
            break
        for p in posts:
            collected[p["post_id"]] = p
        before = min(p["post_id"] for p in posts)
        texted_new = [p for p in collected.values() if p["text"] and p["post_id"] > last_seen]
        if last_seen and before <= last_seen:
            break
        if not last_seen and len(texted_new) >= core.FIRST_RUN_POSTS:
            break
    fresh = sorted((p for p in collected.values() if p["text"] and p["post_id"] > last_seen),
                   key=lambda p: p["post_id"], reverse=True)
    if not last_seen:
        fresh = fresh[:core.FIRST_RUN_POSTS]
    return fresh


def main():
    t0 = time.time()
    channels = core.CHANNELS
    part = load_part()
    save_part(part + 1 if part < 4 else 1)  # следующему запуску — следующая группа (1->2->3->4->1)
    channels = [ch for n, ch in enumerate(channels) if n % 4 == part - 1]
    print(f"=== {datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC запуск: "
          f"группа {part} из 4, каналов: {len(channels)} ===", flush=True)

    state = load_state()
    pending = load_pending()
    print(f"Состояние: {len(state)} каналов, ждут отправки: {len(pending)}", flush=True)

    session = requests.Session()
    session.headers["User-Agent"] = UA
    made = [0]
    fresh_by_channel = {}
    try:
        for ch in channels:
            last_seen = state.get(ch, 0)
            fresh = fetch_channel(session, ch, last_seen, made)
            # у канала без истории первый прогон наберёт ~100 старых постов: их копим молча,
            # чтобы не заваливать бот старьём (уведомятся только посты следующих прогонов)
            first_run = not last_seen
            if fresh:
                insert_posts([{"post_url": p["url"], "channel": p["channel"], "post_id": p["post_id"],
                               "date": p["date"], "text": p["text"], "links": p["links"],
                               "buttons": p.get("buttons", []),
                               "fingerprint": fingerprint_of_post(p),
                               "notified": first_run} for p in fresh])
                save_state(ch, max(last_seen, max(p["post_id"] for p in fresh)))
            print(f"== {ch}: новых постов с текстом {len(fresh)}"
                  + (" (первый прогон, копим молча)" if first_run and fresh else ""), flush=True)
            if not first_run:
                fresh_by_channel[ch] = fresh
    except Exception as e:
        msg = f"⚠️ Парсер вакансий (облако): ошибка при чтении каналов ({type(e).__name__}). Запуск остановлен."
        print(msg, "\n", e, flush=True)
        tg_send(msg)
        sys.exit(1)
    print(f"Запросов к t.me за запуск: {made[0]}", flush=True)

    # Кандидаты: новые посты этого прогона + неотправленные прошлых прогонов
    cands = {p["url"]: p for ch in fresh_by_channel for p in fresh_by_channel[ch]}
    for p in pending:
        cands.setdefault(p["post_url"], {"url": p["post_url"], "channel": p["channel"],
                                         "post_id": p["post_id"], "date": p["date"], "text": p["text"],
                                         "links": p["links"] or [], "buttons": p["buttons"] or []})
    if not cands:
        print("Новых постов нет — нечего обрабатывать.", flush=True)
        return

    # Отпечатки: совпал со старым постом (не из этого прогона) — дубль, молча помечаем
    fps = {url: fingerprint_of_post(p) for url, p in cands.items()}
    dup_urls = {url for url, fp in fps.items() if fp in known_fingerprints(set(fps.values()), set(fps))}

    # Один и тот же пост в разных каналах: оставляем самый ранний (jaccard по шинглам, как локально)
    kept = []
    for url in sorted(cands, key=lambda u: cands[u]["date"] or ""):
        sh = shingles(norm_words(clean_text(core.strip_footer(cands[url]["text"]))[0]))
        if any(jaccard(sh, prev_sh) >= 0.6 for prev_sh in kept):
            dup_urls.add(url)
        else:
            kept.append(sh)

    # Рекламные ссылки каналов — только для каналов, у которых есть что обрабатывать
    boiler = {ch: boilerplate_links(ch) for ch in {cands[u]["channel"] for u in cands}}
    stats = Counter()
    records, done_urls = [], []
    for url, p in cands.items():
        done_urls.append(url)
        if url in dup_urls:
            stats["дубли"] += 1
            continue
        rec = build_record(p, boiler.get(p["channel"], set()), stats)
        if rec:
            records.append(rec)

    sent = notify_new([r for r in records if r["score"] >= core.NOTIFY_MIN_SCORE]) if records else []
    ok = set(sent) | {r["url"] for r in records if r["score"] < core.NOTIFY_MIN_SCORE}
    # Посты, которые не удалось отправить (сеть упала), оставляем pending для повтора.
    failed = {r["url"] for r in records if r["score"] >= core.NOTIFY_MIN_SCORE} - set(sent)
    mark_notified([u for u in done_urls if u not in failed])

    print(f"Обработано: {len(done_urls)}; записей подошло: {len(records)}; "
          f"уведомлено: {len(sent)}; статистика отсева: {dict(stats)}; "
          f"время: {time.time() - t0:.0f} с", flush=True)


if __name__ == "__main__":
    main()
