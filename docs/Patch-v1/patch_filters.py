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
