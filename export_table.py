"""Шаг 3. Убирает дубли и выгружает вакансии в xlsx и csv.

Берёт data/raw_*.json (тексты) и data/labels.jsonl (разметка нейросети).
Результат: vakansii.xlsx, vakansii.csv, data/duplicates.csv (что признано дублем).
Дубль = очень похожий текст (доля общих 4-словных фрагментов от 0,6) ИЛИ одинаковые
должность и компания. Остаётся самый ранний пост, остальные пишутся в колонку «Также в».
"""
import csv
import json
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
DATA = BASE / "data"
CHANNELS = ["onlinevakansii", "jobs_start", "marketing_jobs"]
COLUMNS = ["Дата", "Канал", "Ссылка на пост", "Должность", "Компания", "Зарплата",
           "Формат", "Город", "Контакт", "Уверенность", "Требования", "Также в"]
PERIODS = {"month": "в месяц", "day": "в день", "hour": "в час", "shift": "за смену",
           "project": "за проект"}


def norm_words(text):
    text = re.sub(r"https?://\S+|\[телефон\]|\[email\]|@\w+", " ", text.lower())
    return re.findall(r"[a-zа-яё0-9]+", text)


def shingles(words, n=4):
    return {tuple(words[i:i + n]) for i in range(max(len(words) - n + 1, 1))}


def jaccard(a, b):
    return len(a & b) / len(a | b) if a and b else 0.0


def salary_text(lab):
    lo, hi = lab.get("salary_from"), lab.get("salary_to")
    if lo is None and hi is None:
        return ""
    cur = lab.get("salary_currency") or ""
    per = PERIODS.get(lab.get("salary_period") or "", "")
    parts = []
    if lo is not None:
        parts.append(f"от {lo:,}".replace(",", " "))
    if hi is not None:
        parts.append(f"до {hi:,}".replace(",", " "))
    return " ".join(parts + [cur, per]).strip()


def main():
    raw = {}
    for ch in CHANNELS:
        for p in json.loads((DATA / f"raw_{ch}.json").read_text(encoding="utf-8")):
            raw[p["url"]] = p
    labels = [json.loads(l) for l in (DATA / "labels.jsonl").read_text(encoding="utf-8").splitlines()]
    total = len(labels)
    failed = [r for r in labels if not r["label"]]
    vacancies = [r for r in labels if r["label"] and r["label"].get("is_vacancy")]
    vacancies.sort(key=lambda r: r["date"])  # старые первыми: оставляем самый ранний пост

    kept, dups = [], []
    for r in vacancies:
        words = norm_words(raw[r["url"]]["text"])
        sh = shingles(words)
        key = ((r["label"].get("title") or "").lower().strip(),
               (r["label"].get("company") or "").lower().strip())
        match = None
        for k in kept:
            same_key = key[0] and key[1] and key == k["key"]
            if same_key or jaccard(sh, k["sh"]) >= 0.6:
                match = k
                break
        if match:
            match["also"].append(r["url"])
            dups.append((r["url"], match["rec"]["url"]))
        else:
            kept.append({"rec": r, "sh": sh, "key": key, "also": []})

    rows = []
    for k in sorted(kept, key=lambda k: k["rec"]["date"], reverse=True):
        r, lab = k["rec"], k["rec"]["label"]
        if r["contact"]:
            contact = r["contact"]
        elif r["had_phone"] or r["had_email"]:
            contact = "см. пост (в нём телефон/e-mail)"
        else:
            contact = ""
        rows.append([
            r["date"][:10], r["channel"], r["url"], lab.get("title") or "", lab.get("company") or "",
            salary_text(lab), lab.get("work_format") or "", lab.get("city") or "", contact,
            lab.get("confidence"), lab.get("requirements") or "", "\n".join(k["also"]),
        ])

    with (BASE / "vakansii.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(COLUMNS)
        w.writerows(rows)
    with (DATA / "duplicates.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Дубль (убран)", "Оставлен пост"])
        w.writerows(dups)

    wb = Workbook()
    ws = wb.active
    ws.title = "Вакансии"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        ws.append(row)
    widths = [11, 16, 34, 36, 22, 24, 11, 14, 36, 12, 60, 34]
    for i, wd in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = wd
    for r in ws.iter_rows(min_row=2):
        r[2].hyperlink = r[2].value
        r[2].font = Font(color="0563C1", underline="single")
        for c in r:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(BASE / "vakansii.xlsx")

    print(f"Постов размечено: {total}; JSON не разобран: {len(failed)}")
    print(f"Не вакансии: {total - len(failed) - len(vacancies)}; вакансий: {len(vacancies)}")
    print(f"Дублей убрано: {len(dups)}; в таблице: {len(rows)}")


if __name__ == "__main__":
    main()
