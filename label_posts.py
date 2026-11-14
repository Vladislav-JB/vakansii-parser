"""Шаг 2. Очистка персональных данных и разметка постов нейросетью (kie.ai).

Ключ берётся из переменной среды KIE_API_KEY, нигде не печатается и не сохраняется.
В нейросеть уходит текст БЕЗ телефонов и e-mail. Контакты для отклика (ссылки и @ники)
берутся из ОРИГИНАЛА программой, нейросеть только выбирает номер нужного из списка.

Запуск:
  python label_posts.py --limit 10        # пробный прогон на 10 случайных постах
  python label_posts.py                   # все посты (уже размеченные пропускаются)
  python label_posts.py --budget 30       # остановиться, если потрачено больше 30 кредитов
Результат дописывается в data/labels.jsonl (можно прерывать и запускать снова).
"""
import argparse
import html
import json
import os
import random
import re
import sys
import time
import urllib.request
from pathlib import Path

DATA = Path(__file__).parent / "data"
CHANNELS = ["onlinevakansii", "jobs_start", "marketing_jobs"]
MODEL = "claude-haiku-5-5"
URL = "https://api.kie.ai/claude/v1/messages"

PHONE_RE = re.compile(r"(?<![\w/])(?:\+?[78][\s\-()]*)?(?:\(?\d{3}\)?[\s\-]*)\d{3}[\s\-]*\d{2}[\s\-]*\d{2}(?!\d)")
EMAIL_RE = re.compile(r"[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+")
USERNAME_RE = re.compile(r"(?<![\w@])@([A-Za-z][A-Za-z0-9_]{4,31})\b")
IMG_RE = re.compile(r"\.(?:jpe?g|png|gif|webp|mp4)(?:\?|$)", re.I)


def clean_text(text):
    """Убирает телефоны и e-mail. Возвращает (чистый текст, был_телефон, был_email)."""
    had_email = bool(EMAIL_RE.search(text))
    text = EMAIL_RE.sub("[email]", text)
    had_phone = bool(PHONE_RE.search(text))
    text = PHONE_RE.sub("[телефон]", text)
    return text, had_phone, had_email


def contact_candidates(post):
    """Ссылки и @ники из ОРИГИНАЛА (без телефонов и e-mail)."""
    cands = []
    for link in post["links"]:
        link = html.unescape(link)
        if IMG_RE.search(link) or link in cands:
            continue
        cands.append(link)
    no_email = EMAIL_RE.sub(" ", post["text"])
    for name in USERNAME_RE.findall(no_email):
        handle = "@" + name
        if handle not in cands:
            cands.append(handle)
    return cands[:12]


PROMPT = """Ты помощник HR-агентства. Тебе дан пост из публичного Telegram-канала. Определи, это объявление о вакансии или нет, и верни ТОЛЬКО один JSON-объект без пояснений и без markdown.

Правила:
- Реклама курсов, каналов, ботов, «заработок на опросах», набор в чат и т.п. без конкретной должности и работодателя: is_vacancy=false. Резюме соискателей: is_vacancy=false.
- Ничего не выдумывай. Если данных в тексте нет, ставь null. Зарплату бери только если она явно написана. «По договорённости» или «высокий доход» без цифр: salary_from=null, salary_to=null.
- salary_from и salary_to: числа без пробелов (например 80000). «от 70 000» это salary_from=70000, salary_to=null. «до 8500» это salary_to=8500. salary_currency: RUB, USD, EUR и т.д. salary_period: month, day, hour, shift, project или null.
- work_format: одно из "офис", "удалёнка", "гибрид" или null. «Частичная удалёнка» это гибрид.
- requirements: 1-2 коротких предложения о главных требованиях (до 200 символов) или null.
- contact_index: номер (начиная с 0) из списка КОНТАКТЫ ДЛЯ ОТКЛИКА, куда нужно писать для отклика, или null, если подходящего там нет.
- several_vacancies=true, если в посте несколько разных вакансий (тогда опиши первую).
- confidence: число от 0 до 1, насколько ты уверен в разметке.

Формат ответа:
{"is_vacancy": true, "title": "", "company": null, "salary_from": null, "salary_to": null, "salary_currency": null, "salary_period": null, "work_format": null, "city": null, "requirements": null, "contact_index": null, "several_vacancies": false, "confidence": 0.0}

КОНТАКТЫ ДЛЯ ОТКЛИКА:
%s

ТЕКСТ ПОСТА:
%s"""


def load_posts():
    posts = []
    for ch in CHANNELS:
        posts += json.loads((DATA / f"raw_{ch}.json").read_text(encoding="utf-8"))
    return posts


def call_model(key, prompt):
    body = {"model": MODEL, "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 900, "stream": False}
    req = urllib.request.Request(
        URL, data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": "Bearer " + key, "Content-Type": "application/json",
                 "User-Agent": "Mozilla/5.0"})
    for attempt in range(3):
        try:
            resp = json.load(urllib.request.urlopen(req, timeout=90))
            return resp
        except Exception as e:  # повтор при сетевой ошибке
            print("  ошибка запроса:", type(e).__name__, "- повтор", attempt + 1, flush=True)
            time.sleep(5)
    return None


def parse_json(text):
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="сколько случайных постов (0 = все)")
    ap.add_argument("--budget", type=float, default=30.0, help="максимум кредитов за запуск")
    args = ap.parse_args()

    key = os.environ.get("KIE_API_KEY")
    if not key:
        sys.exit("Нет переменной среды KIE_API_KEY")

    posts = load_posts()
    if args.limit:
        random.Random(7).shuffle(posts)
        posts = posts[:args.limit]

    out_path = DATA / "labels.jsonl"
    done = set()
    if out_path.exists():
        for line in out_path.read_text(encoding="utf-8").splitlines():
            done.add(json.loads(line)["url"])

    spent = 0.0
    n = 0
    with out_path.open("a", encoding="utf-8") as out:
        for post in posts:
            if post["url"] in done:
                continue
            cleaned, had_phone, had_email = clean_text(post["text"])
            cands = contact_candidates(post)
            cand_txt = "\n".join(f"{i}: {c}" for i, c in enumerate(cands)) or "(нет)"
            resp = call_model(key, PROMPT % (cand_txt, cleaned))
            if resp is None:
                print("Остановка: нейросеть не отвечает.")
                break
            spent += float(resp.get("credits_consumed") or 0)
            raw_answer = "".join(c.get("text", "") for c in resp.get("content", []))
            label = parse_json(raw_answer)
            idx = label.get("contact_index") if label else None
            contact = cands[idx] if isinstance(idx, int) and 0 <= idx < len(cands) else None
            rec = {"url": post["url"], "channel": post["channel"], "date": post["date"],
                   "label": label, "contact": contact, "candidates": cands,
                   "had_phone": had_phone, "had_email": had_email,
                   "credits": resp.get("credits_consumed"), "raw_answer": None if label else raw_answer}
            out.write(json.dumps(rec, ensure_ascii=False) + "\n")
            out.flush()
            n += 1
            print(f"{n}: {post['url']} credits={resp.get('credits_consumed')} всего={spent:.2f}"
                  f"{'' if label else '  !!! JSON не разобран'}", flush=True)
            if spent >= args.budget:
                print(f"Остановка: достигнут лимит {args.budget} кредитов.")
                break
            time.sleep(0.7)
    print(f"Размечено за запуск: {n}, потрачено кредитов: {spent:.2f}")


if __name__ == "__main__":
    main()
