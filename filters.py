import re
from datetime import datetime


def filter_date(value):
    if not value or str(value).strip() == "":
        now = datetime.now()
        return f"/{now.month:02d}/{now.year}"
    try:
        for fmt in ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S"]:
            try:
                dt = datetime.strptime(str(value).strip(), fmt)
                return dt.strftime("%d/%m/%Y")
            except ValueError:
                continue
        if isinstance(value, (int, float)):
            from datetime import timedelta
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=float(value))
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(value)


def filter_date_long(value):
    if not value or str(value).strip() == "":
        now = datetime.now()
        return f"tháng {now.month:02d} năm {now.year}"
    try:
        for fmt in ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S"]:
            try:
                dt = datetime.strptime(str(value).strip(), fmt)
                day = dt.day
                month = dt.month
                year = dt.year
                return f"ngày {day:02d} tháng {month:02d} năm {year}"
            except ValueError:
                continue
        if isinstance(value, (int, float)):
            from datetime import timedelta
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=float(value))
            return f"ngày {dt.day:02d} tháng {dt.month:02d} năm {dt.year}"
    except Exception:
        pass
    return str(value)


def filter_number(value):
    if value is None:
        return "0"
    try:
        num = float(re.sub(r"[^\d.,\-]", "", str(value).replace(".", "").replace(",", ".")))
        return f"{num:,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        raise ValueError(f"Cannot format as number: {value}")
