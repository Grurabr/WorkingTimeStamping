from collections import defaultdict
from datetime import datetime

def parse_datetime(date_str):
    try:
        return datetime.strptime(date_str.strip(), "%d.%m.%Y %H.%M.%S")
    except:
        return None
def process_and_group(data):
    grouped = defaultdict(list)

    for row in data:
        person = row.get("Person_number", "").strip()
        grouped[person].append(row)


    for person in grouped:
        grouped[person] = sorted(
            grouped[person],
            key=lambda x: parse_datetime(x.get("Timestamp_date", "")) or datetime.min
        )

    return dict(grouped)
