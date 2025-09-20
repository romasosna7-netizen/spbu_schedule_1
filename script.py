# script.py
# Автоматическая выгрузка расписания СПбГУ → .ics
# Работает по неделям, собирает несколько вперёд и делает единый календарь

import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from datetime import datetime, timedelta, date
from urllib.parse import urljoin
from uuid import uuid4

GROUP_ID = "427997"   # ID твоей группы (если другой — меняй)
WEEKS_AHEAD = 16      # сколько недель вперёд собирать
SITE_ROOT = "https://timetable.spbu.ru"
OUT_ICS = "schedule.ics"

MONTHS = {
    'января':1, 'февраля':2, 'марта':3, 'апреля':4, 'мая':5, 'июня':6,
    'июля':7, 'августа':8, 'сентября':9, 'октября':10, 'ноября':11, 'декабря':12
}

def get_this_monday(d: date):
    return d - timedelta(days=d.weekday())

def find_excel_link_from_page(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = a['href']
        txt = (a.get_text() or "").lower()
        if ".xlsx" in href or "downloadexcel" in href or "скачать" in txt:
            return urljoin(base_url, href)
    return None

def parse_week_excel_bytes(content_bytes, assumed_year=None):
    raw = pd.read_excel(content_bytes, header=None)
    year = assumed_year or datetime.utcnow().year
    df = pd.read_excel(content_bytes, skiprows=3)
    df = df.rename(columns={df.columns[0]:'DayAndDate'})
    df['DayAndDate'] = df['DayAndDate'].ffill()

    events = []
    for _, row in df.iterrows():
        daydate = row.get('DayAndDate')
        time_raw = row.get('Время', None)
        name = row.get('Название', None)
        if pd.isna(time_raw) or pd.isna(name):
            continue

        dd = str(daydate).split()[-2:]
        if len(dd) < 2: 
            continue
        day, mon_name = dd
        if not day.isdigit() or mon_name.lower() not in MONTHS:
            continue
        mon = MONTHS[mon_name.lower()]
        dt_date = datetime(year, mon, int(day)).date()

        time_text = str(time_raw).strip()
        parts = re.split(r'[–—\-]', time_text)
        if len(parts) < 2:
            continue
        start_dt = datetime.strptime(f"{dt_date} {parts[0].strip()}", "%Y-%m-%d %H:%M")
        end_dt   = datetime.strptime(f"{dt_date} {parts[1].strip()}", "%Y-%m-%d %H:%M")

        events.append({
            "start": start_dt,
            "end": end_dt,
            "summary": str(name),
            "location": "" if pd.isna(row.get('Места проведения')) else str(row.get('Места проведения')),
            "teacher": "" if pd.isna(row.get('Преподаватели')) else str(row.get('Преподаватели'))
        })
    return events

def events_to_ics(events):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//spbu-schedule//EN",
        "CALSCALE:GREGORIAN"
    ]
    now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    for e in events:
        uid = str(uuid4())
        def dtf(dt): return dt.strftime("%Y%m%dT%H%M%S")
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now}",
            f"DTSTART:{dtf(e['start'])}",
            f"DTEND:{dtf(e['end'])}",
            f"SUMMARY:{e['summary']}"
        ]
        if e['location']:
            lines.append(f"LOCATION:{e['location']}")
        if e['teacher']:
            lines.append(f"DESCRIPTION:Преподаватель: {e['teacher']}")
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\n".join(lines)

def main():
    today = datetime.utcnow().date()
    start_monday = get_this_monday(today)
    all_events = []

    for w in range(WEEKS_AHEAD):
        week_date = start_monday + timedelta(weeks=w)
        page_url = f"{SITE_ROOT}/EARTH/StudentGroupEvents/Primary/{GROUP_ID}/{week_date}"
        r = requests.get(page_url)
        if r.status_code != 200:
            continue
        excel_link = find_excel_link_from_page(r.text, page_url)
        if not excel_link:
            continue
        rx = requests.get(excel_link)
        if rx.status_code == 200:
            evs = parse_week_excel_bytes(rx.content)
            all_events.extend(evs)
    if not all_events:
        print("Нет событий")
        return

    ics_text = events_to_ics(all_events)
    with open(OUT_ICS, "w", encoding="utf-8") as f:
        f.write(ics_text)
    print("Файл schedule.ics создан")

if name == "__main__":
    main()
