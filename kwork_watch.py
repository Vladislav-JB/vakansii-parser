"""Облачная версия сторожа заказов Kwork: GitHub Actions + Supabase.

Перенос локального kwork-notify-watch.ps1 (правила те же, формат уведомлений —
утверждённый навыком kwork-notify):
  - читает открытые страницы kwork.ru/projects (категория 37, страницы 1-2);
  - фильтры: бюджет от 1 000 ₽, откликов не больше 10, сфера «сайты/боты/лендинги»,
    ориентир 5 000 ₽ для простой визитки (в разы дороже — пропуск), стоп-сферы
    (WordPress, Bitrix, парсеры, личные кабинеты, оплаты и т.п.);
  - учёт просмотренных и отправленных — таблица kw_orders в Supabase
    (локальные kwork-seen.json/kwork-pending.json недоступны из облака);
  - уведомление — напрямую в бот через Bot API; недоставленное доотправится
    в следующем прогоне (notified=false);
  - будущий адрес демо https://web-atelie.ru/demo-<id>/ закрепляется в таблице;
    репозиторий и заглушка НЕ создаются — только по команде пользователя.
Отклики заказчикам не отправляются вообще.

Переменные окружения: те же секреты, что у парсера вакансий
(SUPABASE_URL, SUPABASE_SERVICE_KEY, TG_BOT_TOKEN, TG_CHAT_ID).
"""
import json
import re
import sys
import time

import requests

from cloud_monitor import sb_request, tg_send  # те же секреты и REST-помощники

PAGES = (1, 2)
EXCLUDED = {"3270786", "3271296"}  # нежелательные заказы по решению пользователя
BATCH = 200

# PostgREST требует, чтобы у всех объектов в одной вставке был одинаковый набор
# полей (иначе PGRST102 «All object keys must match»). Отсеянные и подходящие
# заказы заполняют разные поля, поэтому набор выравнивается перед записью.
ORDER_KEYS = ("order_id", "name", "price", "price_max", "offers",
              "slug", "url", "reason", "notified")


class KworkUnavailable(Exception):
    """Kwork отдал не список заказов (антибот, лимит, сбой) — это не поломка кода."""

OUTSIDE_RE = re.compile(
    r"wordpress|вордпресс|\bwp\b|битрикс|bitrix|opencart|joomla|modx|агрегатор|парсер|парсинг|"
    r"личн\w* кабинет|онлайн[ -]?оплат|интеграц\w*\s+(?:с\s+)?1[сc]|электронн\w* дневник|lms|saas|"
    r"пиксель в пиксель|правк\w* существующ|доработ\w* сайт|перенос\w* сайт|перенести сайт|только с отзывами",
    re.I)
MUST_RE = re.compile(r"сайт|лендинг|лэндинг|визитк|landing|одностранич|бот|ботов|телеграм|ocr|aiogram|cron", re.I)
SIMPLE_RE = re.compile(r"визитк|прост\w*\s+(?:сайт|лендинг|лэндинг)|небольш\w*\s+сайт", re.I)
EXTRA_RE = re.compile(r"интеграц|crm|срм|многостранич|уникальн\w* дизайн|индивидуальн\w* дизайн|"
                      r"анимаци|калькулятор|каталог", re.I)


def clean(value):
    """Без тегов и HTML-сущностей, как Clean в ps1."""
    text = re.sub(r"<[^>]+>", " ", str(value or ""))
    import html as html_lib
    return re.sub(r"\s+", " ", html_lib.unescape(text)).strip()


def reject_reason(order):
    """Те же правила отбора, что в kwork-notify-watch.ps1 (RejectReason)."""
    oid = str(order.get("id", ""))
    if oid in EXCLUDED:
        return "excluded-by-user"
    try:
        price = float(order["priceLimit"])
    except (KeyError, TypeError, ValueError):
        return "unknown-budget"
    if price < 1000:
        return "budget-below-1000"
    try:
        offers = int(order["kwork_count"])
    except (KeyError, TypeError, ValueError):
        return "unknown-offers"
    if offers < 0:
        return "unknown-offers"
    if offers > 10:
        return "offers-above-10"
    text = clean(order.get("name")) + " " + clean(order.get("description"))
    if OUTSIDE_RE.search(text):
        return "outside-current-scope"
    if not MUST_RE.search(text):
        return "not-a-small-site"
    if SIMPLE_RE.search(text) and not EXTRA_RE.search(text) and price >= 15000:
        return "simple-site-budget-outlier"
    return None


def future_slug(order_id):
    """Свободно ли имя demo-<id> на GitHub (только проверка, ничего не создаём)."""
    slug = f"demo-{order_id}"
    try:
        r = requests.head(f"https://api.github.com/repos/Vladislav-JB/{slug}",
                          headers={"User-Agent": "WebAtelie-Kwork-Notify"}, timeout=20)
        if r.status_code == 404:
            return slug
        if r.status_code == 200:
            return f"demo-{order_id}-2"  # имя занято — берём запасное
    except requests.RequestException as e:
        print(f"GitHub не ответил для {slug}: {type(e).__name__}", flush=True)
        raise
    return slug


def fetch_orders(session):
    """Открытые заказы категории 37 (сайты), страницы 1-2. С повторами.

    Kwork иногда отвечает HTML-заглушкой (антибот, лимит) вместо JSON — это не
    ошибка кода: пробуем ещё раз, а если не вышло, поднимаем KworkUnavailable,
    чтобы прогон завершился без красного креста в Actions.
    """
    orders = []
    for page in PAGES:
        body = None
        for attempt in range(3):
            resp = None
            try:
                resp = session.post("https://kwork.ru/projects", data={"c": "37", "page": str(page)},
                                    headers={"X-Requested-With": "XMLHttpRequest",
                                             "User-Agent": "Mozilla/5.0 WebAtelie-Kwork-Notify"},
                                    timeout=30)
            except requests.RequestException as e:
                print(f"    стр.{page}: сеть дёрнулась ({type(e).__name__})", flush=True)
            if resp is not None and resp.status_code == 200:
                try:
                    body = resp.json()
                except ValueError:
                    body = None
            if body is not None:
                break
            if attempt + 1 < 3:
                wait = 10 * (attempt + 1)
                got = "не JSON" if resp is not None and resp.status_code == 200 else (
                    f"HTTP {resp.status_code}" if resp is not None else "нет ответа")
                print(f"    стр.{page}: вместо заказов {got}, повтор {attempt + 2} из 3 через {wait} с",
                      flush=True)
                time.sleep(wait)
        else:
            raise KworkUnavailable(f"страница {page}: за 3 попытки заказов не получили")
        data = (body.get("data") or {}).get("pagination") or {}
        items = [x for x in (data.get("data") or []) if x]
        orders += items
        print(f"стр.{page}: HTTP 200, заказов {len(items)}", flush=True)
    return {str(o["id"]): o for o in orders}


def insert_orders(rows):
    """Пишет новые заказы в kw_orders. У всех строк — один и тот же набор полей."""
    padded = [{k: r.get(k) for k in ORDER_KEYS} for r in rows]
    for i in range(0, len(padded), BATCH):
        sb_request("POST", "kw_orders", prefer="resolution=ignore-duplicates,return=minimal",
                   body=padded[i:i + BATCH])


def load_seen():
    rows = sb_request("GET", "kw_orders", params={"select": "order_id", "order_id": "not.is.null",
                                                 "limit": 50000}).json()
    return {r["order_id"] for r in rows}


def load_pending():
    """Заказы прошлых прогонов, уведомление которых не удалось доставить."""
    return sb_request("GET", "kw_orders", params={
        "select": "order_id,name,price,price_max,offers,slug,url",
        "notified": "eq.false", "reason": "is.null", "limit": 100}).json()


def mark_notified(ids):
    from datetime import datetime, timezone
    for i in range(0, len(ids), BATCH):
        sb_request("PATCH", "kw_orders", params={"order_id": f"in.({','.join(ids[i:i + BATCH])})"},
                   body={"notified": True, "notified_at": datetime.now(timezone.utc).isoformat()})


def order_text(order_id, name, price, price_max, offers, slug):
    """Уведомление по утверждённому образцу навыка (короткое, с будущим адресом)."""
    url = f"https://web-atelie.ru/{slug}/"
    budget = f"Бюджет: {int(price)} ₽" + (f" (до {int(price_max)} ₽)" if price_max > price else "") + f" · откликов: {offers}"
    return (f"Kwork: {name}\n{budget}\nПокупатель: статистика не проверена\n"
            f"https://kwork.ru/projects/{order_id}/view\n\n"
            f"Адрес будущего демо: {url}\nПока не опубликован. После «делай заказ {order_id}» "
            f"здесь появится заглушка, затем сайт.\n\n"
            f"✍️ Текст отклика (после готовности демо; сначала напишите «делай заказ {order_id}»):\n"
            f"Здравствуйте! Мне уже есть что показать: собрал сайт по вашему заданию.\n"
            f"Это предварительный вариант, внесу любые изменения, по вашему желанию!\n\n"
            f"Можно посмотреть здесь: {url}\n\n"
            f"Другие примеры сайтов, которые сделаны в Веб-Ателье: https://web-atelie.ru/\n\n"
            f"Будущий адрес, пока не опубликован.")


def notify_order(row):
    """Шлёт уведомление одного заказа. Возвращает True, если Telegram подтвердил."""
    return tg_send(order_text(row["order_id"], row["name"], row["price"] or 0,
                              row["price_max"] or 0, row["offers"] or 0, row["slug"]))


def main():
    session = requests.Session()
    try:
        orders = fetch_orders(session)
    except KworkUnavailable as e:
        # Kwork ответил заглушкой вместо заказов — не падаем (иначе письмо о сбое),
        # просто сообщаем и ждём следующего запуска через 5 минут
        msg = f"Kwork не отдал заказы ({e}). Следующая попытка через 5 минут."
        print("⏸ " + msg, flush=True)
        tg_send("⏸ Сторож Kwork: " + msg)
        return
    except Exception as e:
        msg = f"⚠️ Сторож Kwork (облако): ошибка при чтении заказов ({type(e).__name__}). Прогон остановлен."
        print(msg, "\n", e, flush=True)
        tg_send(msg)
        sys.exit(1)

    seen = load_seen()
    pending = load_pending()
    fresh = [oid for oid in orders if oid.isdigit() and oid not in seen]
    print(f"Заказов получено: {len(orders)}, новых: {len(fresh)}, ждут доставки: {len(pending)}", flush=True)

    # Новые заказы: прогоняем фильтры и пишем в базу (отсеянные — сразу notified)
    rows = []
    eligible = []
    for oid in fresh:
        o = orders[oid]
        reason = reject_reason(o)
        try:
            price = float(o["priceLimit"])
        except (KeyError, TypeError, ValueError):
            price = None
        try:
            price_max = float(o.get("possiblePriceLimit") or 0)
        except (TypeError, ValueError):
            price_max = 0
        try:
            offers = int(o["kwork_count"])
        except (KeyError, TypeError, ValueError):
            offers = None
        if reason:
            rows.append({"order_id": oid, "name": clean(o.get("name"))[:200], "price": price,
                         "price_max": price_max, "offers": offers, "reason": reason, "notified": True})
            print(f"skip {oid} {reason}", flush=True)
        else:
            slug = future_slug(oid)
            row = {"order_id": oid, "name": clean(o.get("name"))[:200], "price": price,
                   "price_max": price_max, "offers": offers, "slug": slug,
                   "url": f"https://kwork.ru/projects/{oid}/view",
                   "notified": False}
            rows.append(row)
            eligible.append(row)
    insert_orders(rows)

    # К доставке: подходящие этого прогона + недоставленные прошлых прогонов
    todo = {r["order_id"]: {"order_id": r["order_id"], "name": r["name"], "price": r["price"],
                            "price_max": r["price_max"], "offers": r["offers"], "slug": r["slug"]}
            for r in eligible}
    for p in pending:
        todo.setdefault(p["order_id"], p)
    sent = [oid for oid, row in todo.items() if notify_order(row)]
    if sent:
        mark_notified(sent)
    print(f"Уведомлений отправлено: {len(sent)} из {len(todo)}", flush=True)


if __name__ == "__main__":
    main()
