"""Бесплатный мониторинг вакансий и заказов в публичных Telegram-каналах (без нейросети).

Что делает за один запуск:
  1. Читает новые посты каналов через веб-версию t.me/s/<канал> (без аккаунта, только чтение,
     пауза 8-12 секунд между запросами; при любом ответе не 200 сразу останавливается).
     При первом запуске берёт примерно последние FIRST_RUN_POSTS постов канала,
     при следующих запусках только те, что новее уже прочитанных (data/it_state.json).
     Все прочитанные посты копятся в data/it_posts.json.
  2. Убирает из текста телефоны и e-mail.
  3. Оставляет посты, где есть слова по нашим темам (ИИ, сайты, боты и автоматизация, внедрение).
     Подборки из десятков вакансий в одном посте пропускаются (в них нет ссылок на каждую вакансию).
  4. Убирает дубли (один пост в разных каналах).
  5. Сохраняет накопленный список в vakansii_it.xlsx и vakansii_it.csv.

Запуск:  python it_monitor.py
Без сети (пересобрать таблицу из накопленных постов):  python it_monitor.py --offline
"""
import atexit
import csv
import html as html_lib
import json
import random
import re
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from export_table import jaccard, norm_words, shingles
from fetch_posts import UA, parse_page
from label_posts import EMAIL_RE, IMG_RE, USERNAME_RE, clean_text

BASE = Path(__file__).parent
DATA = BASE / "data"
STATE = DATA / "it_state.json"
POSTS = DATA / "it_posts.json"
LOCK = DATA / "it_run.lock"

def load_channels():
    """Читает channels.txt: имена каналов, @имена или ссылки t.me/..., по одному на строку."""
    path = BASE / "channels.txt"
    if not path.exists():
        sys.exit("Нет файла channels.txt со списком каналов (по одному на строку).")
    channels, seen = [], set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        name = re.sub(r"^(?:https?://)?(?:t\.me|telegram\.me)/(?:s/)?", "", line, flags=re.I)
        name = name.split("?")[0].split("/")[0].lstrip("@").strip()
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{4,31}", name):
            print(f"channels.txt: пропускаю строку «{line}» (не похоже на имя публичного канала)")
            continue
        if name.lower() not in seen:
            seen.add(name.lower())
            channels.append(name)
    if not channels:
        sys.exit("В channels.txt нет ни одного канала.")
    return channels


CHANNELS = load_channels()

# Разбиение на две группы для чередующихся запусков каждые 10 минут:
#   --part 1 — каналы с чётными номерами в списке, --part 2 — с нечётными.
# Без --part (например, run_notify.bat руками) обходится весь список сразу.
PART = None
if "--part" in sys.argv:
    _i = sys.argv.index("--part")
    if _i + 1 < len(sys.argv):
        PART = sys.argv[_i + 1]
if PART in ("1", "2"):
    CHANNELS = [ch for _n, ch in enumerate(CHANNELS) if _n % 2 == (0 if PART == "1" else 1)]
FIRST_RUN_POSTS = 100      # сколько постов с текстом брать с канала при первом запуске
MAX_PAGES = 8              # предохранитель: не больше 8 страниц (по 20 постов) на канал за запуск
PAUSE = (8, 12)            # пауза между запросами, секунды
BOILERPLATE_MIN = 5        # ссылка, встречающаяся в стольких постах канала, считается рекламой канала

# Темы и слова (регулярные выражения, поиск без учёта регистра)
TOPICS = {
    "ИИ": [
        r"(?<![а-яёa-z])ии(?![а-яёa-z])", r"(?<![a-z])ai(?![a-z])", r"(?<![a-z])llm", r"нейросет",
        r"нейронн", r"chat\s?gpt", r"(?<![a-z])gpt", r"искусственн\w+ интеллект", r"промпт",
        r"prompt", r"машинн\w+ обучени", r"(?<![a-z])ml(?![a-z])", r"(?<![a-z])rag(?![a-z])",
        r"langchain", r"генеративн", r"компьютерн\w+ зрени|computer vision", r"openai|claude|gemini",
    ],
    "Сайты": [
        r"сайт", r"лендинг|landing", r"tilda|тильд", r"wordpress|вордпресс", r"вёрстк|верстк",
        r"веб[- ]?(?:разработ|дизайн)|web[- ]?(?:dev|разработ|design)", r"frontend|front-end|фронтенд",
        r"битрикс|bitrix", r"интернет-магазин", r"ux/ui|ui/ux",
    ],
    "Боты и автоматизация": [
        r"(?<![а-яё])бот(?:ы|ов|а|у)?(?![а-яё])", r"chat-?bot|чат-?бот", r"автоматизац",
        r"n8n|make\.com|zapier|no-?code|low-?code", r"парсинг|парсер", r"(?<![a-z])rpa(?![a-z])",
    ],
    "Внедрение": [
        r"внедрени", r"(?<![a-z])crm(?![a-z])", r"amo\s?crm|амо\s?crm", r"1[сc](?![а-яa-z0-9])",
        r"(?<![a-z])erp(?![a-z])",
    ],
}
TOPIC_RE = {t: [re.compile(p, re.I) for p in pats] for t, pats in TOPICS.items()}
GENERIC_TITLES = re.compile(r"^(?:новая вакансия|вакансия|срочно|ищем|требуется|hot|удалённо|удаленно)$", re.I)
COLUMNS = ["Дата", "Канал", "Ссылка на пост", "Ссылка для отклика", "Заголовок", "Оценка 0-10", "Теги",
           "Тема", "Совпавшие слова", "Описание", "Также в", "Новое в этом запуске"]

# кнопки под постом: «КОНТАКТЫ», «Откликнуться» дают ссылку для отклика, «ИСПОЛНИТЕЛЬ НАЙДЕН» = закрыто
CONTACT_BUTTON_RE = re.compile(r"контакт|отклик|связаться|подать заявку|написать|ставк|bid", re.I)
CLOSED_BUTTON_RE = re.compile(r"исполнитель найден|заказ закрыт|вакансия закрыта|набор закрыт", re.I)

# Заказы Kwork: число откликов в посте не указано, а читать страницы Kwork скриптом нельзя
# (блокировка по IP). Поэтому смотрим на возраст поста: старше этого срока заказ считаем
# «уже с кучей откликов» и убираем на лист «Отсеяно». Поменяй число часов под себя.
KWORK_MAX_AGE_HOURS = 24
# Заказы Kwork сюда не нужны: их ловит отдельный сторож. True = показывать их в таблице
# (тогда действует правило про возраст выше), False = убирать все на лист «Отсеяно».
INCLUDE_KWORK = False
KWORK_ORDER_RE = re.compile(r"kwork\.ru/projects/\d+", re.I)

# --- Стоп-слова: пост с ними отсеивается сразу (попадает на лист «Отсеяно» с причиной) ---
STOP_WORDS = [
    r"курсов(?:ая|ую|ой|ые)\b", r"дипломн", r"реферат", r"заработок на ", r"быстрый заработок",
    r"пассивн\w+ доход", r"без вложений", r"сетевой маркетинг|(?<![a-z])mlm(?![a-z])",
    r"ставки на спорт|казино|беттинг", r"простые опросы|прохождени\w+ опрос|на опросах",
    r"исполнитель найден", r"розыгрыш", r"инфобизнес", r"заработал\w* [\d ]+\s*(?:₽|руб)", r"интенсив|вебинар|мастер-класс",
]
STOP_RE = [re.compile(p, re.I) for p in STOP_WORDS]
# в начале текста: резюме, предложения услуг, реклама каналов и курсов (проверка по итогам 30 строк)
HEAD_STOP_RE = re.compile(
    r"#резюме|#помогу|ищу работу|ищу заказы|ищу проект|предлагаю услуги|помогаю (?:решать|вам|компани|бизнес)|"
    r"подборк\w+ (?:тг-)?каналов|в одном боте|ламповое комьюнити|освоение искусственного интеллекта|"
    r"фриланс-бот нашел|могу быть полезен|привет! я |теперь делаю |беру заказы|принимаю заказы|берусь за", re.I)
# в заголовке: роли не про разработку и внедрение (сильный штраф, даже если в заголовке есть «AI»)
NON_IT_TITLE_RE = re.compile(
    r"smm|креатор|креатив|контент|копирайтер|маркетолог|affiliate|sales|продаж|дизайнер|designer|монтаж|"
    r"ассистент|менеджер по|партн[её]рств", re.I)
TITLE_NON_IT_PENALTY = 12
MAX_AGE_DAYS = 60  # посты старше этого срока не показываем (лист «Отсеяно», причина «старый пост»)

# --- Баллы (идея взята из scan-agent: группы слов с весом и штрафами) ---
# Каждое слово из группы считается один раз на пост. Итог = сумма / 4, от 0 до 10.
SCORING = {
    "ядро": (10, [
        r"нейросет", r"(?<![a-z])llm", r"chat\s?gpt", r"искусственн\w+ интеллект", r"промпт|prompt",
        r"(?<![а-яёa-z])ии(?![а-яёa-z])|(?<![a-z])ai(?![a-z])",
        r"лендинг|landing", r"tilda|тильд", r"wordpress|вордпресс", r"разработк\w+ сайт|создани\w+ сайт",
        r"(?<![а-яё])бот(?:ы|ов|а|у)?(?![а-яё])|чат-?бот", r"n8n|make\.com|zapier",
        r"внедрени\w+ (?:ии|ai|нейросет|crm)",
    ]),
    "связанное": (5, [
        r"автоматизац",
        r"(?<![a-z])ml(?![a-z])", r"(?<![a-z])rag(?![a-z])", r"langchain", r"openai|claude|gemini",
        r"(?<![a-z])crm(?![a-z])|amo\s?crm|битрикс|bitrix", r"1[сc](?![а-яa-z0-9])",
        r"python", r"(?<![a-z])api(?![a-z])|интеграц", r"парсинг|парсер", r"вёрстк|верстк|frontend|фронтенд",
        r"сайт", r"интернет-магазин", r"no-?code|low-?code", r"разработ(?:ать|ка|чик)|доработ",
    ]),
    "дополнительное": (2, [
        r"react", r"typescript", r"node\.?js", r"django|fastapi|flask", r"php|laravel",
        r"vue", r"docker", r"postgres|sql", r"telegram|телеграм",
    ]),
    "штраф: не IT-роль": (-6, [
        r"менеджер по продажам|sales manager|account executive|business development",
        r"рекрутер|recruit|talent acquisition|(?<![a-z])hr(?![a-z])", r"бухгалтер|юрист|финансов\w+ контролл",
        r"smm|копирайтер|контент-менеджер|маркетолог|marketing manager",
        r"дизайнер|designer|монтажёр|монтажер|сценарист|директолог|таргетолог|креатор|креативщик",
    ]),
    "штраф: высокий уровень": (-3, [
        r"(?<![a-z])(?:head|chief|vp|director)(?![a-z])|директор|руководитель отдела",
    ]),
}
SCORING_RE = {g: (w, [re.compile(p, re.I) for p in pats]) for g, (w, pats) in SCORING.items()}
TITLE_BONUS = 6  # если слово из «ядра» есть в заголовке
MAX_HITS_PER_GROUP = 2  # больше двух слов из одной группы не добавляют баллов

# --- Теги (идея из scan-agent): быстрый бесплатный разбор по ключевым словам ---
TAG_RULES = {
    "AI/ML": [r"нейросет", r"нейронн", r"(?<![a-z])llm", r"chat\s?gpt", r"openai", r"машинн\w+ обучени",
              r"(?<![a-z])ml(?![a-z])", r"(?<![а-яёa-z])ии(?![а-яёa-z])|(?<![a-z])ai(?![a-z])", r"промпт|prompt"],
    "Telegram": [r"telegram|телеграм"],
    "Бот": [r"(?<![а-яё])бот(?:ы|ов|а|у)?(?![а-яё])", r"чат-?бот|chat-?bot"],
    "Автоматизация": [r"автоматизац", r"n8n|make\.com|zapier", r"(?<![a-z])rpa(?![a-z])"],
    "Парсинг": [r"парсинг|парсер|scraping|playwright|puppeteer|selenium"],
    "Лендинг/сайт": [r"лендинг|landing", r"сайт", r"интернет-магазин"],
    "No-code": [r"tilda|тильд", r"no-?code|low-?code", r"n8n|make\.com|zapier"],
    "WordPress/Bitrix": [r"wordpress|вордпресс", r"битрикс|bitrix"],
    "CRM/1С": [r"(?<![a-z])crm(?![a-z])|amo\s?crm", r"1[сc](?![а-яa-z0-9])"],
    "Python": [r"python|django|fastapi|flask"],
    "JS/Frontend": [r"react", r"vue", r"typescript", r"node\.?js", r"frontend|front-end|фронтенд"],
    "PHP": [r"php|laravel|symfony"],
    "Mobile": [r"android|ios|flutter|react native|swift|kotlin"],
    "UX/UI": [r"ux/ui|ui/ux|ui-?дизайн|ux-?дизайн|product designer"],
}
TAG_RE = {t: [re.compile(p, re.I) for p in pats] for t, pats in TAG_RULES.items()}
# похоже на страницу вакансии или формы отклика (выбирается раньше прочих ссылок)
APPLY_RE = re.compile(r"vacanc|career|/jobs?[/.]|jobs\.|hh\.ru|geekjob|lnkd\.in|linkedin\.com/jobs|apply|"
                      r"forms?\.|typeform|notion\.|ashby|lever\.co|greenhouse|workable|teamtailor|"
                      r"peopleforce|huntflow|getmatch|habr\.com/.*career", re.I)


def find_topics(text):
    """Возвращает {тема: [совпавшие слова]}."""
    found = {}
    for topic, regs in TOPIC_RE.items():
        words = []
        for rg in regs:
            for m in rg.finditer(text):
                w = m.group(0).strip().lower()
                if w and w not in words:
                    words.append(w)
        if words:
            found[topic] = words
    return found


def is_digest(text):
    """Подборка: много строк-рубрик «📍…» (в них нет ссылок на отдельные вакансии)."""
    if text.count("➡️") >= 3:  # несколько вакансий подряд, у каждой своя стрелка с контактом
        return True
    first = next((l for l in text.splitlines() if l.strip()), "")
    if re.search(r"дайджест|подборка (?:вакансий|компаний)", first, re.I):
        return True
    return sum(1 for l in text.splitlines() if l.strip().startswith("📍")) >= 3


def make_title(text):
    for line in text.splitlines():
        if line.strip().startswith("📍"):      # рубрика («📍Backend»), не заголовок
            continue
        line = re.sub(r"^[^\w#]+", "", line.strip())
        line = re.sub(r"^(?:#\S+\s*)+", "", line).strip()
        line = re.sub(r"^(?:вакансия|требуется|ищем)\s*[:\-–—]\s*", "", line, flags=re.I)
        if len(line) >= 4 and not GENERIC_TITLES.match(line):
            return line[:140]
    return ""


def strip_ref(url):
    """Убирает реферальный параметр ref=... (каналы вроде freelance_zakazy дописывают свой)."""
    from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
    parts = urlsplit(url)
    query = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True) if k != "ref"]
    return urlunsplit(parts._replace(query=urlencode(query)))


def reply_link(post, boilerplate):
    """Ссылка для отклика из ОРИГИНАЛА поста (рекламные ссылки канала и приглашения не берём)."""
    own = re.compile(r"^https?://(?:t\.me|telegram\.me)/(?:%s)(?:/|$)" % "|".join(CHANNELS), re.I)
    for name, href in post.get("buttons", []):  # кнопка под постом («КОНТАКТЫ», «Откликнуться»)
        href = strip_ref(html_lib.unescape(href))
        if CONTACT_BUTTON_RE.search(name) and not own.match(href):
            return href, []
    links = []
    for l in post["links"]:
        l = html_lib.unescape(l)
        if (IMG_RE.search(l) or own.match(l) or "t.me/addlist" in l or "t.me/+" in l
                or "t.me/joinchat" in l or l in boilerplate or l in links):
            continue
        links.append(l)
    for name in USERNAME_RE.findall(EMAIL_RE.sub(" ", post["text"])):
        handle = "https://t.me/" + name
        if handle not in links and handle not in boilerplate and not own.match(handle):
            links.append(handle)
    human = [l for l in links if re.match(r"^https?://(?:t\.me|telegram\.me)/[A-Za-z]\w{4,}$", l)
             and "bot" not in l.lower()]
    apply_like = [l for l in links if APPLY_RE.search(l) and l not in human]
    best = human[0] if human else (apply_like[0] if apply_like else (links[0] if links else ""))
    return best, [l for l in links if l != best]


def get_with_retry(session, url, attempts=3):
    """GET с повторами: разовый обрыв сети не должен убивать весь обход.
    Итоговое уведомление «ошибка при чтении каналов» шлётся только если не помогли все повторы."""
    last = None
    for a in range(attempts):
        try:
            return session.get(url, timeout=30)
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
            last = e
            if a + 1 < attempts:
                wait = 8 * (a + 1)
                print(f"    сеть дёрнулась ({type(e).__name__}), повтор {a + 2} из {attempts} через {wait} с", flush=True)
                time.sleep(wait)
    raise last


def fetch_new(state, store):
    """Читает новые посты, дописывает их в store. Возвращает множество новых ссылок на посты."""
    session = requests.Session()
    session.headers["User-Agent"] = UA
    new_urls, made = set(), 0
    for ch in CHANNELS:
        last_seen = state.get(ch, 0)
        collected, before = {}, None
        for page in range(MAX_PAGES):
            if made:
                time.sleep(random.uniform(*PAUSE))
            url = f"https://t.me/s/{ch}" + (f"?before={before}" if before else "")
            resp = get_with_retry(session, url)
            made += 1
            print(f"{ch} стр.{page + 1}: HTTP {resp.status_code}", flush=True)
            if resp.status_code != 200:
                print("Остановка: сервер ответил не 200. Больше ничего не запрашиваю.")
                notify(f"⚠️ Парсер вакансий остановлен: Telegram ответил кодом {resp.status_code} на канал {ch}. "
                       "Больше ничего не запрашиваю. Лучше подождать несколько часов.")
                sys.exit(1)
            posts = parse_page(resp.text, ch)
            if not posts:
                break
            for p in posts:
                collected[p["post_id"]] = p
            before = min(p["post_id"] for p in posts)
            texted_new = [p for p in collected.values() if p["text"] and p["post_id"] > last_seen]
            if last_seen and before <= last_seen:
                break
            if not last_seen and len(texted_new) >= FIRST_RUN_POSTS:
                break
        fresh = sorted((p for p in collected.values() if p["text"] and p["post_id"] > last_seen),
                       key=lambda p: p["post_id"], reverse=True)
        if not last_seen:
            fresh = fresh[:FIRST_RUN_POSTS]
        if collected:
            state[ch] = max(state.get(ch, 0), max(collected))
        for p in fresh:
            store[p["url"]] = p
            new_urls.add(p["url"])
        print(f"== {ch}: новых постов с текстом {len(fresh)}", flush=True)
    print(f"Запросов к t.me за запуск: {made}")
    return new_urls


def strip_footer(text):
    """Отрезает подпись канала после строки-разделителя (➖➖➖…): в ней реклама и ссылки на ботов."""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.fullmatch(r"\s*[➖─—_=]{5,}\s*", line):
            return "\n".join(lines[:i]).strip() or text
    return text


def stop_reason(text):
    """Причина отсева по стоп-словам или None."""
    head = HEAD_STOP_RE.search(text[:300])
    if head:
        return f"не вакансия (резюме, услуги или реклама): {head.group(0).strip().lower()}"
    for rg in STOP_RE:
        m = rg.search(text)
        if m:
            return f"стоп-слово: {m.group(0).strip().lower()}"
    return None


def score_post(title, text):
    """Оценка 0-10 по группам слов (идея из scan-agent). Возвращает (оценка, пояснение)."""
    raw, notes = 0, []
    for group, (weight, regs) in SCORING_RE.items():
        hits = sorted({m.group(0).strip().lower() for rg in regs for m in [rg.search(text)] if m})
        if hits:
            counted = min(len(hits), MAX_HITS_PER_GROUP)  # длинный пост не должен выигрывать только длиной
            raw += weight * counted
            notes.append(f"{group} {weight:+d}x{counted}: {', '.join(hits)}")
    if any(rg.search(title) for rg in SCORING_RE["ядро"][1]):
        raw += TITLE_BONUS
        notes.append(f"слово из ядра в заголовке +{TITLE_BONUS}")
    if NON_IT_TITLE_RE.search(title):
        raw -= TITLE_NON_IT_PENALTY
        notes.append(f"в заголовке роль не про разработку -{TITLE_NON_IT_PENALTY}")
    return min(10, max(0, round(raw / 4))), "; ".join(notes)


def tags_of(text):
    return [t for t, regs in TAG_RE.items() if any(rg.search(text) for rg in regs)]


def build_records(store, new_urls):
    """Из всех накопленных постов делает записи для подходящих. Возвращает (записи, статистика)."""
    link_count = Counter()
    for p in store.values():
        for l in set(html_lib.unescape(x) for x in p["links"]):
            link_count[(p["channel"], l)] += 1
    stats = {"всего": len(store), "подборок": 0, "без тем": 0, "стоп-слова": 0, "подходит": 0,
             "старые Kwork": 0}
    records, excluded = [], []
    for p in store.values():
        if p["channel"] not in CHANNELS:  # канал убрали из channels.txt: его посты не показываем
            continue
        if is_digest(p["text"]):
            stats["подборок"] += 1
            continue
        body = strip_footer(p["text"])  # без подписи канала с рекламой
        topics = find_topics(body)
        if not topics:
            stats["без тем"] += 1
            continue
        text, had_phone, had_email = clean_text(body)
        title = make_title(text)
        reason = stop_reason(text)
        if p["date"] and (datetime.now(timezone.utc) - datetime.fromisoformat(p["date"])).days > MAX_AGE_DAYS:
            reason = f"старый пост (больше {MAX_AGE_DAYS} дней)"
        for name, _href in p.get("buttons", []):
            if CLOSED_BUTTON_RE.search(name):
                reason = f"закрыто: кнопка «{name}» (на момент чтения)"
        if reason:
            stats["стоп-слова"] += 1
            excluded.append([p["date"][:10], p["channel"], p["url"], title, reason])
            continue
        stats["подходит"] += 1
        score, score_notes = score_post(title, text)
        boiler = {l for (ch, l), n in link_count.items() if ch == p["channel"] and n >= BOILERPLATE_MIN}
        best, _ = reply_link(p, boiler)
        if KWORK_ORDER_RE.search(best):
            reason = None
            if not INCLUDE_KWORK:
                reason = "заказ Kwork: их ловит отдельный сторож"
            elif p["date"]:
                age_h = (datetime.now(timezone.utc) - datetime.fromisoformat(p["date"])).total_seconds() / 3600
                if age_h > KWORK_MAX_AGE_HOURS:
                    reason = f"заказ Kwork старше {KWORK_MAX_AGE_HOURS} ч: скорее всего уже много откликов"
            if reason:
                stats["подходит"] -= 1
                stats["старые Kwork"] += 1
                excluded.append([p["date"][:10], p["channel"], p["url"], title, reason])
                continue
        if not best:
            best = "в посте нет ссылки (см. пост; там телефон/e-mail)" if (had_phone or had_email) \
                else "в посте нет ссылки (см. пост)"
        words = sorted({w for ws in topics.values() for w in ws})
        records.append({
            "date": p["date"], "channel": p["channel"], "url": p["url"], "reply": best, "title": title,
            "topics": list(topics), "words": words, "text": text, "score": score,
            "score_notes": score_notes, "tags": tags_of(text), "new": p["url"] in new_urls,
        })
    return records, excluded, stats


def dedupe(records):
    """Один и тот же пост в разных каналах оставляем один раз (самый ранний)."""
    kept = []
    for r in sorted(records, key=lambda r: r["date"]):
        sh = shingles(norm_words(r["text"]))
        for k in kept:
            if jaccard(sh, k["sh"]) >= 0.6:
                k["also"].append(r["url"])
                break
        else:
            kept.append({"rec": r, "sh": sh, "also": []})
    return kept


def write_tables(records, excluded):
    kept = dedupe(records)
    # сначала свежие, затем (устойчивая сортировка) по оценке: выше оценка - выше в таблице
    kept.sort(key=lambda k: k["rec"]["date"], reverse=True)
    kept.sort(key=lambda k: k["rec"]["score"], reverse=True)
    rows = []
    for k in kept:
        r = k["rec"]
        rows.append([r["date"][:10], r["channel"], r["url"], r["reply"], r["title"], r["score"],
                     ", ".join(r["tags"]), ", ".join(r["topics"]), r["score_notes"], r["text"][:3000],
                     "\n".join(k["also"]), "да" if r["new"] else ""])
    with (BASE / "vakansii_it.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(COLUMNS)
        w.writerows(rows)
    wb = Workbook()

    def fill_sheet(ws, sheet_rows):
        ws.append(COLUMNS)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="305496")
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for row in sheet_rows:
            ws.append(row)
        for i, wd in enumerate([11, 16, 34, 38, 44, 10, 26, 20, 40, 90, 34, 12], 1):
            ws.column_dimensions[get_column_letter(i)].width = wd
        for r in ws.iter_rows(min_row=2):
            for idx in (2, 3):  # ссылки кликабельные
                val = r[idx].value
                if isinstance(val, str) and val.startswith("http"):
                    r[idx].hyperlink = val
                    r[idx].font = Font(color="0563C1", underline="single")
            for c in r:
                c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    best_ws = wb.active
    best_ws.title = "Лучшие (оценка 7+)"
    fill_sheet(best_ws, [r for r in rows if r[5] >= 7])
    fill_sheet(wb.create_sheet("Все подходящие"), rows)
    ws2 = wb.create_sheet("Отсеяно")
    ws2.append(["Дата", "Канал", "Ссылка на пост", "Заголовок", "Причина"])
    for c in ws2[1]:
        c.font = Font(bold=True)
    for row in sorted(excluded, reverse=True):
        ws2.append(row)
    for i, wd in enumerate([11, 16, 40, 60, 32], 1):
        ws2.column_dimensions[get_column_letter(i)].width = wd
    try:
        wb.save(BASE / "vakansii_it.xlsx")
    except PermissionError:  # файл открыт в Excel
        wb.save(BASE / "vakansii_it_новый.xlsx")
        print("ВНИМАНИЕ: vakansii_it.xlsx открыт в Excel, сохранил в vakansii_it_новый.xlsx. "
              "Закрой Excel и запусти ещё раз.")
    return rows, len(records) - len(kept)


TG_SCRIPT = Path.home() / ".claude" / "vibe-tools" / "tg.ps1"  # отправка сообщения тебе в Telegram
NOTIFY_MIN_SCORE = 7      # в уведомление попадают новые строки с такой оценкой и выше
NOTIFY_MAX_ITEMS = 15     # в одном запуске присылаем не больше стольких строк
NOTIFY = "--notify" in sys.argv or "--notify-test" in sys.argv


def notify(text):
    """Отправляет текст в Telegram через tg.ps1 (только при запуске с --notify). Токен не выводится."""
    if not NOTIFY:
        return
    if not TG_SCRIPT.exists():
        print("Уведомление не отправлено: нет", TG_SCRIPT)
        return
    import subprocess
    tmp = DATA / "_notify.txt"
    tmp.write_text(text, encoding="utf-8")
    res = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(TG_SCRIPT),
                          "-File", str(tmp)], capture_output=True, text=True, timeout=120)
    if res.returncode != 0 or "sent" not in res.stdout:
        print("Не удалось отправить уведомление в Telegram (подробности в логе tg.ps1 не выводятся).")


def notify_new(rows, title="🔎 Новые вакансии и заказы (оценка 7+)"):
    """Шлёт новые строки таблицы в Telegram, порциями до ~3500 символов."""
    items = []
    for i, r in enumerate(rows[:NOTIFY_MAX_ITEMS], 1):
        date, channel, url, reply, ttl, score, tags = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
        desc = re.sub(r"\s+", " ", r[9]).strip()[:200]
        items.append(f"{i}. {ttl}\nОценка {score} · {channel} · {date}" + (f"\nТеги: {tags}" if tags else "")
                     + f"\nОтклик: {reply}\nПост: {url}\n{desc}…")
    chunk, header = "", f"{title}: {len(rows)}\n\n"
    for it in items:
        if len(chunk) + len(it) > 3500:
            notify(header + chunk)
            chunk, header = "", "(продолжение)\n\n"
        chunk += it + "\n\n"
    if chunk:
        if len(rows) > NOTIFY_MAX_ITEMS:
            chunk += f"…и ещё {len(rows) - NOTIFY_MAX_ITEMS}, полный список в таблице vakansii_it.xlsx"
        notify(header + chunk)


def acquire_lock():
    """Замок: чтобы плановый запуск не наезжал на ещё работающий предыдущий.
    Замок старше 30 минут считается забытым (например, после жёсткого выключения) и не мешает."""
    DATA.mkdir(exist_ok=True)
    if LOCK.exists():
        try:
            age = time.time() - LOCK.stat().st_mtime
        except OSError:
            age = 0
        if age < 1800:
            print(f"Другой прогон ещё работает (замок держится {int(age)} с). Этот запуск пропускаю.", flush=True)
            sys.exit(0)
    LOCK.write_text(datetime.now().isoformat(), encoding="utf-8")


def release_lock():
    try:
        LOCK.unlink()
    except OSError:
        pass


def resolve_part():
    """--part auto: группа выбирается после замка, чтобы чередовали только реально
    работающие прогоны (пропущенные по замку не забирают ход). Чередование через
    файл-счётчик data/it_part.txt: 1 -> 2 -> 1 -> ..."""
    global PART
    if PART != "auto":
        return
    _flag = DATA / "it_part.txt"
    try:
        _last = _flag.read_text(encoding="utf-8").strip()
    except OSError:
        _last = ""
    PART = "1" if _last not in ("1", "2") else ("2" if _last == "1" else "1")
    try:
        _flag.write_text(PART, encoding="utf-8")
    except OSError:
        pass


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")  # чтобы русские буквы не ломали запуск из Планировщика
    DATA.mkdir(exist_ok=True)
    acquire_lock()
    atexit.register(release_lock)
    resolve_part()
    global CHANNELS
    if PART in ("1", "2"):
        CHANNELS = [ch for _n, ch in enumerate(CHANNELS) if _n % 2 == (0 if PART == "1" else 1)]
    label = f"группа {PART} из 2, каналов: {len(CHANNELS)}" if PART else f"все каналы: {len(CHANNELS)}"
    print(f"=== {datetime.now():%Y-%m-%d %H:%M} запуск ({'--notify' if NOTIFY else 'без уведомлений'}; {label}) ===", flush=True)
    state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}
    store = json.loads(POSTS.read_text(encoding="utf-8")) if POSTS.exists() else {}
    new_urls = set()
    if "--offline" not in sys.argv:
        try:
            new_urls = fetch_new(state, store)
        except SystemExit:
            raise
        except Exception as e:  # сеть упала и т.п.: сообщаем и выходим
            notify(f"⚠️ Парсер вакансий: ошибка при чтении каналов ({type(e).__name__}). Запуск остановлен.")
            raise
        POSTS.write_text(json.dumps(store, ensure_ascii=False, indent=1), encoding="utf-8")
        STATE.write_text(json.dumps(state, indent=1), encoding="utf-8")
    records, excluded, stats = build_records(store, new_urls)
    rows, dups = write_tables(records, excluded)
    if "--notify-test" in sys.argv:  # пробное уведомление: 3 лучшие строки таблицы, чтобы увидеть формат
        notify_new([r for r in rows if r[5] >= NOTIFY_MIN_SCORE][:3], title="🧪 ТЕСТ уведомления парсера вакансий")
    elif NOTIFY and "--offline" not in sys.argv:
        fresh = [r for r in rows if r[11] == "да" and r[5] >= NOTIFY_MIN_SCORE]
        if fresh:
            notify_new(fresh)
        print(f"Уведомление: новых строк с оценкой {NOTIFY_MIN_SCORE}+ : {len(fresh)}")
    scores = [r[5] for r in rows]
    print(f"Постов в накопленной базе: {stats['всего']}; подборок пропущено: {stats['подборок']}; "
          f"без нужных слов: {stats['без тем']}; отсеяно стоп-словами: {stats['стоп-слова']}; "
          f"заказов Kwork убрано: {stats['старые Kwork']}; подошло: {stats['подходит']}")
    print(f"Дублей убрано: {dups}; в таблице: {len(rows)} (оценка 7-10: {sum(s >= 7 for s in scores)}, "
          f"4-6: {sum(4 <= s <= 6 for s in scores)}, 0-3: {sum(s <= 3 for s in scores)})")
    print("Файлы: vakansii_it.xlsx, vakansii_it.csv")


if __name__ == "__main__":
    main()
