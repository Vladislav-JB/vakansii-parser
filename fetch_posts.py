"""Шаг 1. Читает публичные Telegram-каналы через веб-версию t.me/s/<канал>.

Без входа в аккаунт, только чтение. Запросы редкие: пауза 8-12 секунд.
При любой ошибке (не 200, 429) сразу останавливается, чтобы не получить блокировку IP.
Результат: data/raw_<канал>.json
"""
import json
import random
import sys
import time
from pathlib import Path

import requests
from lxml import html as lh

CHANNELS = ["onlinevakansii", "jobs_start", "marketing_jobs"]
TARGET = 100          # сколько постов с текстом нужно с каждого канала
MAX_PAGES = 12        # защита от бесконечного цикла
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
OUT = Path(__file__).parent / "data"


def parse_page(page_html, channel):
    doc = lh.fromstring(page_html)
    posts = []
    for wrap in doc.xpath('//div[contains(@class,"tgme_widget_message_wrap")]'):
        msg = wrap.xpath('.//div[contains(@class,"tgme_widget_message ")]')
        if not msg:
            continue
        data_post = msg[0].get("data-post", "")
        if "/" not in data_post:
            continue
        post_id = int(data_post.split("/")[1])
        date = wrap.xpath('.//a[contains(@class,"tgme_widget_message_date")]/time/@datetime')
        text_nodes = wrap.xpath('.//div[contains(@class,"tgme_widget_message_text")]')
        text, links = "", []
        if text_nodes:
            node = text_nodes[0]
            for br in node.xpath(".//br"):
                br.tail = "\n" + (br.tail or "")
            for a in node.xpath(".//a"):
                href = a.get("href", "")
                if href and not href.startswith("?q=") and "t.me/s/" not in href:
                    links.append(href)
            text = node.text_content().strip()
        buttons = []  # кнопки под постом: [название, ссылка] (например «КОНТАКТЫ»)
        for a in wrap.xpath('.//a[contains(@class,"tgme_widget_message_inline_button")]'):
            if a.get("href"):
                buttons.append([a.text_content().strip(), a.get("href")])
        posts.append({
            "buttons": buttons,
            "channel": channel,
            "post_id": post_id,
            "url": f"https://t.me/{channel}/{post_id}",
            "date": date[0] if date else "",
            "text": text,
            "links": sorted(set(links)),
        })
    return posts


def main():
    OUT.mkdir(exist_ok=True)
    session = requests.Session()
    session.headers["User-Agent"] = UA
    requests_made = 0
    for channel in CHANNELS:
        collected = {}
        before = None
        for page in range(MAX_PAGES):
            url = f"https://t.me/s/{channel}" + (f"?before={before}" if before else "")
            if requests_made:
                time.sleep(random.uniform(8, 12))
            resp = session.get(url, timeout=30)
            requests_made += 1
            print(f"{channel} стр.{page + 1}: HTTP {resp.status_code}", flush=True)
            if resp.status_code != 200:
                print("Остановка: сервер ответил не 200. Ничего больше не запрашиваю.")
                sys.exit(1)
            posts = parse_page(resp.text, channel)
            if not posts:
                break
            for p in posts:
                collected[p["post_id"]] = p
            before = min(p["post_id"] for p in posts)
            with_text = [p for p in collected.values() if p["text"]]
            if len(with_text) >= TARGET:
                break
        result = sorted((p for p in collected.values() if p["text"]),
                        key=lambda p: p["post_id"], reverse=True)[:TARGET]
        (OUT / f"raw_{channel}.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
        dates = [p["date"][:10] for p in result if p["date"]]
        print(f"== {channel}: сохранено {len(result)} постов с текстом, "
              f"даты {min(dates) if dates else '?'} .. {max(dates) if dates else '?'}", flush=True)
    print(f"Всего запросов к t.me: {requests_made}")


if __name__ == "__main__":
    main()
