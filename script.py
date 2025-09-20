import requests
import pandas as pd
from datetime import datetime, timedelta, date
from uuid import uuid4
from openpyxl import load_workbook


# === Настройки ===
GROUP_ID = "427997"
WEEKS_AHEAD = 4
SITE_ROOT = "https://timetable.spbu.ru"
OUT_ICS = "schedule.ics"

# === Функции ===
def get_this_monday(d: date):
    return d - timedelta(days=d.weekday())


def parse_week(start_date: date):
    url = f"{SITE_ROOT}/StudentGroupEvents/ExcelWeek?studentGroupId={GROUP_ID}&weekMonday={start_date:%Y-%m-%d}"
    print("[xlsx url]", url)
    r = requests.get(url)
    if r.status_code != 200:
        print("[warn] нет файла по адресу", url)
        return []

    with open("tmp.xlsx", "wb") as f:
        f.write(r.content)

    # Читаем Excel через Pandas как строки
    df = pd.read_excel("tmp.xlsx", header=None, skiprows=4, dtype=str)
    df = df.fillna('')  # заменяем NaN на пустые строки

    events = []

    for _, row in df.iterrows():
        dt_raw = row[0].strip()
        time_raw = row[1].strip()
        subj = row[2].strip()
        room = row[3].strip()
        # преподаватель в календарь не добавляем

        # пропускаем пустые строки
        if not dt_raw or not time_raw or not subj:
            continue

        # Парсим дату, локализуем русский месяц
        try:
            # заменяем день недели и приводим месяц к английскому для pd.to_datetime
            import re
            dt_text = re.sub(r'^\w+\s+', '', dt_raw)  # убираем день недели
            # для корректного распознавания русских месяцев можно через mapping
            month_map = {
                "января":"January","февраля":"February","марта":"March",
                "апреля":"April","мая":"May","июня":"June","июля":"July",
                "августа":"August","сентября":"September","октября":"October",
                "ноября":"November","декабря":"December"
            }
            for ru, en in month_map.items():
                dt_text = dt_text.replace(ru, en)
            dt = pd.to_datetime(dt_text, dayfirst=True, errors='coerce')
            if pd.isna(dt):
                continue
        except:
            continue

        # Парсим диапазон времени
        time_raw = time_raw.replace("–", "-")  # заменяем длинное тире
        if "-" not in time_raw:
            continue
        start_str, end_str = [t.strip() for t in time_raw.split("-")]
        try:
            start_time = datetime.strptime(start_str, "%H:%M").time()
            end_time = datetime.strptime(end_str, "%H:%M").time()
        except:
            continue

        start_dt = datetime.combine(dt.date(), start_time)
        end_dt = datetime.combine(dt.date(), end_time)

        events.append({
            "uid": str(uuid4()),
            "summary": subj,
            "location": room,
            "dtstart": start_dt,
            "dtend": end_dt,
        })
        print("[event]", subj, start_dt, "-", end_dt)

    return events


def make_ics(events):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
    ]
    for ev in events:
        lines += [
            "BEGIN:VEVENT",
            f"UID:{ev['uid']}",
            f"SUMMARY:{ev['summary']}",
            f"DTSTART:{ev['dtstart'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{ev['dtend'].strftime('%Y%m%dT%H%M%S')}",
            f"LOCATION:{ev['location']}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")

    with open(OUT_ICS, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("[ok] файл сохранён:", OUT_ICS)

def main():
    today = date.today()
    monday = get_this_monday(today)

    all_events = []
    for i in range(WEEKS_AHEAD):
        week_start = monday + timedelta(days=i*7)
        print("=== Обработка недели:", week_start, "===")
        all_events.extend(parse_week(week_start))

    print("[total events]", len(all_events))
    if all_events:
        make_ics(all_events)
    else:
        print("[warn] событий нет, файл не создан")

if __name__ == "__main__":
    main()
