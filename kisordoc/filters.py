"""
PATCH: filters.py
BUG FIX: filter_number — logic replace bị ngược gây ValueError với input dạng "1,500,000"
"""

import re
from datetime import datetime


def filter_date(value):
    if not value or str(value).strip() == "":
        now = datetime.now()
        return f"/{now.month:02d}/{now.year}"
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    try:
        for fmt in [
            "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", 
            "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"
        ]:
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
    val_str = str(value) if value is not None else ""
    if "/" in val_str:
        parts = val_str.split("/")
        if len(parts) == 3:
            day_part = parts[0]
            month_part = parts[1]
            year_part = parts[2]
            if not day_part.strip() or not month_part.strip():
                day_str = f" {day_part.strip()} " if day_part.strip() else "   "
                month_str = f" {month_part.strip()} " if month_part.strip() else "   "
                year_str = f" {year_part.strip()}" if year_part.strip() else year_part
                return f"ngày{day_str}tháng{month_str}năm{year_str}"

    if not value or str(value).strip() == "":
        now = datetime.now()
        return f"tháng {now.month:02d} năm {now.year}"
    if isinstance(value, datetime):
        return f"ngày {value.day:02d} tháng {value.month:02d} năm {value.year}"
    try:
        for fmt in [
            "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", 
            "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"
        ]:
            try:
                dt = datetime.strptime(str(value).strip(), fmt)
                return f"ngày {dt.day:02d} tháng {dt.month:02d} năm {dt.year}"
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
    """
    FIX: Detect format trước, không replace mù.
    - Input VN: "1.500.000" (dấu chấm = nghìn)  → 1500000 → "1.500.000"
    - Input US: "1,500,000" (dấu phẩy = nghìn)   → 1500000 → "1.500.000"
    - Input float: 1500000.0                       → 1500000 → "1.500.000"
    - Input int: 1500000                           → "1.500.000"
    """
    if value is None:
        return "0"
    # Nếu đã là số
    if isinstance(value, (int, float)):
        import math
        if math.isnan(value):
            return "0"
        return f"{int(round(value)):,}".replace(",", ".")

    s = str(value).strip()
    if not s:
        return "0"

    # Bỏ tất cả ký tự không phải số, dấu phẩy, chấm, trừ
    s_clean = re.sub(r"[^\d.,\-]", "", s)
    if not s_clean:
        raise ValueError(f"Cannot format as number: {value}")

    # Detect format:
    # Trường hợp 1: "1.500.000" hoặc "1.500" — dấu chấm là nghìn (VN style)
    #   → dấu chấm xuất hiện nhiều lần, hoặc phần sau dấu chấm cuối có 3 chữ số
    # Trường hợp 2: "1,500,000" — dấu phẩy là nghìn (US style)
    # Trường hợp 3: "1500000" — không có dấu phân cách

    dot_count = s_clean.count(".")
    comma_count = s_clean.count(",")

    if dot_count > 1:
        # VN style: "1.500.000" → bỏ chấm
        s_clean = s_clean.replace(".", "")
    elif comma_count > 1:
        # US style: "1,500,000" → bỏ phẩy
        s_clean = s_clean.replace(",", "")
    elif dot_count == 1 and comma_count == 1:
        # Mixed: xác định cái nào là decimal separator
        dot_pos = s_clean.index(".")
        comma_pos = s_clean.index(",")
        if dot_pos < comma_pos:
            # "1.500,00" → chấm là nghìn, phẩy là thập phân
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            # "1,500.00" → phẩy là nghìn, chấm là thập phân
            s_clean = s_clean.replace(",", "")
    elif dot_count == 1:
        # "1500.75" hoặc "1.500"
        parts = s_clean.split(".")
        if len(parts[1]) == 3:
            # Có thể là VN style "1.500" → nghìn
            s_clean = s_clean.replace(".", "")
        # else: là số thập phân "1500.75" → giữ nguyên
    elif comma_count == 1:
        # "1500,75" hoặc "1,500"
        parts = s_clean.split(",")
        if len(parts[1]) == 3:
            # US style nghìn
            s_clean = s_clean.replace(",", "")
        else:
            # Thập phân kiểu EU "1500,75" → đổi thành "1500.75"
            s_clean = s_clean.replace(",", ".")

    try:
        num = float(s_clean)
        return f"{int(round(num)):,}".replace(",", ".")
    except (ValueError, TypeError):
        raise ValueError(f"Cannot format as number: {value}")


def number_to_vietnamese_words(value) -> str:
    if value is None:
        return "Không"
    
    if isinstance(value, (int, float)):
        import math
        if math.isnan(value):
            return "Không"
        val_int = int(round(value))
    else:
        s = str(value).strip()
        if not s:
            return "Không"
        try:
            clean_s = filter_number(s)
            clean_s = clean_s.replace(".", "")
            val_int = int(clean_s)
        except Exception:
            clean_digits = "".join(c for c in s if c.isdigit() or c == "-")
            if not clean_digits or clean_digits == "-":
                return str(value)
            try:
                val_int = int(clean_digits)
            except Exception:
                return str(value)

    if val_int == 0:
        return "Không"

    is_negative = val_int < 0
    val_str = str(abs(val_int))

    pad_len = (3 - len(val_str) % 3) % 3
    val_str = "0" * pad_len + val_str

    groups = [val_str[i:i+3] for i in range(0, len(val_str), 3)]
    num_groups = len(groups)

    units_words = ["", " nghìn", " triệu", " tỷ"]
    units = []
    for idx in range(num_groups):
        unit_idx = idx % 4
        super_idx = idx // 4
        unit_str = units_words[unit_idx]
        if super_idx > 0 and unit_idx == 0:
            unit_str = " tỷ" * (super_idx + 1)
        elif super_idx > 0:
            unit_str = unit_str + " tỷ" * super_idx
        units.append(unit_str)
    units = units[::-1]

    digits_map = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

    def read_three_digits(grp, is_first_group):
        h, t, u = int(grp[0]), int(grp[1]), int(grp[2])
        if h == 0 and t == 0 and u == 0:
            return ""

        words = []
        if not is_first_group or h > 0:
            words.append(digits_map[h] + " trăm")

        if t == 0:
            if u > 0 and (not is_first_group or h > 0):
                words.append("lẻ")
        elif t == 1:
            words.append("mười")
        else:
            words.append(digits_map[t] + " mươi")

        if u > 0:
            if u == 1 and t > 1:
                words.append("mốt")
            elif u == 5 and t > 0:
                words.append("lăm")
            elif u == 4 and t > 1:
                words.append("tư")
            else:
                words.append(digits_map[u])

        return " ".join(words)

    group_words = []
    for idx, grp in enumerate(groups):
        grp_word = read_three_digits(grp, idx == 0)
        if grp_word:
            unit = units[idx]
            group_words.append(grp_word + unit)

    result = " ".join(group_words).strip()
    
    if result:
        result = result[0].upper() + result[1:]

    if is_negative:
        if result:
            result = result[0].lower() + result[1:]
        result = "Âm " + result

    return result


def filter_num2text(value) -> str:
    try:
        return number_to_vietnamese_words(value)
    except Exception:
        return str(value)


def _parse_to_datetime(value) -> datetime:
    if not value or str(value).strip() == "":
        return datetime.now()
    if isinstance(value, datetime):
        return value
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime()
    try:
        for fmt in ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S"]:
            try:
                return datetime.strptime(str(value).strip(), fmt)
            except ValueError:
                continue
        if isinstance(value, (int, float)):
            from datetime import timedelta
            base = datetime(1899, 12, 30)
            return base + timedelta(days=float(value))
    except Exception:
        pass
    raise ValueError(f"Cannot parse value as datetime: {value}")


def filter_day(value) -> str:
    try:
        dt = _parse_to_datetime(value)
        return f"{dt.day:02d}"
    except Exception:
        return ""


def filter_month(value) -> str:
    try:
        dt = _parse_to_datetime(value)
        return f"{dt.month:02d}"
    except Exception:
        return ""


def filter_year(value) -> str:
    try:
        dt = _parse_to_datetime(value)
        return f"{dt.year:04d}"
    except Exception:
        return ""


def filter_add_days(value, days):
    try:
        dt = _parse_to_datetime(value)
        from datetime import timedelta
        return dt + timedelta(days=int(days))
    except Exception:
        return value


def filter_add_months(value, months):
    try:
        dt = _parse_to_datetime(value)
        m_to_add = int(months)
        new_year = dt.year + (dt.month - 1 + m_to_add) // 12
        new_month = (dt.month - 1 + m_to_add) % 12 + 1
        import calendar
        _, last_day = calendar.monthrange(new_year, new_month)
        new_day = min(dt.day, last_day)
        return dt.replace(year=new_year, month=new_month, day=new_day)
    except Exception:
        return value


def filter_date_diff(value, other_value) -> int:
    try:
        dt1 = _parse_to_datetime(value)
        dt2 = _parse_to_datetime(other_value)
        return abs((dt1 - dt2).days)
    except Exception:
        return 0


def filter_quarter(value) -> str:
    try:
        dt = _parse_to_datetime(value)
        quarter_map = ["I", "II", "III", "IV"]
        q = (dt.month - 1) // 3
        return f"Quý {quarter_map[q]}/{dt.year}"
    except Exception:
        return ""


def filter_weekday(value) -> str:
    try:
        dt = _parse_to_datetime(value)
        weekday_map = {
            0: "Thứ Hai",
            1: "Thứ Ba",
            2: "Thứ Tư",
            3: "Thứ Năm",
            4: "Thứ Sáu",
            5: "Thứ Bảy",
            6: "Chủ Nhật"
        }
        return weekday_map[dt.weekday()]
    except Exception:
        return ""


def filter_date_text(value) -> str:
    try:
        dt = _parse_to_datetime(value)
        day_str = number_to_vietnamese_words(dt.day).lower()
        month_str = number_to_vietnamese_words(dt.month).lower()
        year_str = number_to_vietnamese_words(dt.year).lower()
        return f"Ngày {day_str} tháng {month_str} năm {year_str}"
    except Exception:
        return str(value)


