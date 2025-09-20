# script.py
# Автоматическая выгрузка расписания СПбГУ в .ics
# Работает по неделям, собирает несколько вперёд и делает единый календарь

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta, date
from urllib.parse import urljoin
from uuid import uuid4

# === Настройки ===
GROUP_ID = "427997"            # ID твоей группы
WEEKS_AHEAD = 4                # сколько недель вперёд собирать
SITE_ROOT = "https://timetable.spbu.ru"
OUT_ICS = "schedule.ics"

# === Функции ===
def get_this_monday(d: date):
    return d - timedelta(days=d.weekday())

def parse_week(start_date: date):
    url = f"{SITE_ROOT}/EARTH/StudentGroupEvents/Primary/{GROUP_ID}/{start_date:%Y-%m-%d}"
    print("[week url]", url)
    r = requests.get(url)
    r.raise_for_status()

    soup = BeautifulSoup(r.text, "html.parser")
    links = [a["href"] for a in soup.find_all("a", href=True) if a["href"].endswith(".xlsx")]
    if not links:
        print("[warn] нет .xlsx ссылок на неделе", start_date)
        return []

    xlsx_url = urljoin(SITE_ROOT, links[0])
    print("[xlsx]", xlsx_url)

    r = requests.get(xlsx_url)
    r.raise_for_status()

    with open("tmp.xlsx", "wb") as f:
        f.write(r.content)

    df = pd.read_excel("tmp.xlsx", skiprows=3)
    events = []

    for _, row in df.iterrows():
        try:
            subj = str(row["Дисциплина"]).strip()
            dt_str = str(row["Дата"]).strip()
            time_str = str(row["Время"]).strip()
            room = str(row.get("Аудитория", "")).strip()

            if not subj or subj == "nan":
                continue

            dt = pd.to_datetime(dt_str, dayfirst=True)
            start_time = datetime.strptime(time_str.split("-")[0].strip(), "%H:%M").time()
            end_time = datetime.strptime(time_str.split("-")[1].strip(), "%H:%M").time()

            start_dt = datetime.combine(dt.date(), start_time)
            end_dt = datetime.combine(dt.date(), end_time)

            ev = {
                "uid": str(uuid4()),
                "summary": subj,
                "location": room,
                "dtstart": start_dt,
                "dtend": end_dt,
            }
            events.append(ev)
            print("[event]", subj, start_dt, "-", end_dt)
        except Exception as e:
            print("[error row]", e)

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
        all_events.extend(parse_week(week_start))

    print("[total events]", len(all_events))
    if all_events:
        make_ics(all_events)
    else:
        print("[warn] событий нет, файл не создан")

if __name__ == "__main__":
    main()
