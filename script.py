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
    r.raise_for_status()
    with open("tmp.xlsx", "wb") as f:
        f.write(r.content)

    # читаем Excel без заголовков, начиная с 5-й строки
    df = pd.read_excel("tmp.xlsx", header=None, skiprows=4, dtype=str)
    df = df.fillna('')

    events = []

    # очистка даты от лишнего
    def clean_date(dt_raw: str) -> str:
        import re
        dt_text = dt_raw.replace("\n", " ").strip()

        weekday_words = [
            "понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье",
            "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"
        ]
        pattern = r"^(?:" + "|".join(weekday_words) + r")\s+"
        dt_text = re.sub(pattern, "", dt_text, flags=re.IGNORECASE)

        month_map = {
            "января": "January", "февраля": "February", "марта": "March",
            "апреля": "April", "мая": "May", "июня": "June", "июля": "July",
            "августа": "August", "сентября": "September", "октября": "October",
            "ноября": "November", "декабря": "December"
        }
        for ru, en in month_map.items():
            dt_text = dt_text.replace(ru, en)

        return dt_text

    for _, row in df.iterrows():
        dt_raw = row[0].strip()
        time_raw = row[1].strip()
        subj = row[2].strip()
        room = row[3].strip()
        teacher = row[4].strip()

        if not subj or not dt_raw or not time_raw:
            continue

        # парсим дату
        dt_text = clean_date(dt_raw)
        dt = pd.to_datetime(dt_text, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            print("[skip date]", dt_raw)
            continue

        # парсим время
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

        ev = {
            "uid": str(uuid4()),
            "summary": subj,
            "location": room,
            "description": teacher,
            "dtstart": start_dt,
            "dtend": end_dt,
        }
        events.append(ev)
        print("[event]", subj, start_dt, "-", end_dt, "|", teacher)

    return events


def make_ics(events, filename="schedule.ics"):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("BEGIN:VCALENDAR\n")
        f.write("VERSION:2.0\n")
        f.write("PRODID:-//Schedule Parser//EN\n")

        for ev in events:
            f.write("BEGIN:VEVENT\n")
            f.write(f"UID:{ev['uid']}\n")
            f.write(f"SUMMARY:{ev['summary']}\n")
            if ev["location"]:
                f.write(f"LOCATION:{ev['location']}\n")
            if ev["description"]:
                f.write(f"DESCRIPTION:{ev['description']}\n")
            f.write(f"DTSTART:{ev['dtstart'].strftime('%Y%m%dT%H%M%S')}\n")
            f.write(f"DTEND:{ev['dtend'].strftime('%Y%m%dT%H%M%S')}\n")
            f.write("END:VEVENT\n")

        f.write("END:VCALENDAR\n")
    print(f"[ok] файл сохранён: {filename}")

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
