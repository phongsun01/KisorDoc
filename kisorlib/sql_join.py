"""
kisorlib/sql_join.py — Refactor v3.2 (Phase B)
────────────────────────────────────────────────
SQL join helpers, tách ra từ utils.py.

Exports:
  - _OP_MAP, _JOIN_RE
  - validate_sql_identifier
  - parse_join_expression
  - resolve_sheet_query
  - _parse_repeat_sheet_config
"""

import re

_JOIN_RE = re.compile(r'.+\s+(?:<\*>|<\*|\*>|\*)\s+.+\s*@\s*.+', re.DOTALL)

# Thứ tự ưu tiên: <*> trước <* trước *> trước *
_OP_MAP = [
    (" <*>", "FULL OUTER JOIN"),
    (" <*",  "LEFT JOIN"),
    (" *>",  "RIGHT JOIN"),
    (" *",   "INNER JOIN"),
]


def validate_sql_identifier(name: str) -> str:
    if not name:
        return ""
    pattern = r"^[A-Za-z0-9_\s\-\.\#\u00C0-\u1EF9]+$"
    if not re.match(pattern, name):
        raise ValueError(
            f"⚠️ Cảnh báo bảo mật: Phát hiện ký tự không hợp lệ trong tên bảng "
            f"hoặc cột (SQL Injection Risk): '{name}'"
        )
    return name


def parse_join_expression(expr: str) -> str:
    """
    Cú pháp rút gọn cho cột Sheet trong Options:
      Table1 <* Table2 @ key           → LEFT JOIN, cùng tên cột
      Table1 <* Table2 @ key1 = key2   → LEFT JOIN, khác tên cột
      Table1 *> Table2 @ key           → RIGHT JOIN
      Table1 * Table2 @ key            → INNER JOIN
      Table1 <*> Table2 @ key          → FULL OUTER JOIN
      SELECT ...                        → passthrough
    """
    s = expr.strip()
    if s.lower().startswith("select"):
        return s

    # Tách phần điều kiện WHERE nếu có (hỗ trợ điều kiện lọc thêm)
    where_clause = ""
    where_parts = re.split(r'\s+WHERE\s+', s, flags=re.IGNORECASE)
    if len(where_parts) > 1:
        s = where_parts[0].strip()
        where_clause = " WHERE " + where_parts[1].replace("==", "=").strip()

    if "@" not in s:
        return f'SELECT * FROM "{validate_sql_identifier(s)}"{where_clause}'

    join_part, key_raw = s.split("@", 1)
    join_part = join_part.strip()
    key_raw   = key_raw.strip()

    join_type = None
    t1 = t2 = ""
    for sym, jt in _OP_MAP:
        if sym in join_part:
            join_type = jt
            left, right = join_part.split(sym.strip(), 1)
            t1 = validate_sql_identifier(left.strip())
            t2 = validate_sql_identifier(right.strip())
            break

    if not join_type:
        return f'SELECT * FROM "{validate_sql_identifier(s)}"{where_clause}'

    if "=" in key_raw:
        k1, k2 = [validate_sql_identifier(k.strip()) for k in key_raw.split("=", 1)]
    else:
        k1 = k2 = validate_sql_identifier(key_raw)

    return (
        f'SELECT * FROM "{t1}" {join_type} "{t2}" '
        f'ON "{t1}"."{k1}" = "{t2}"."{k2}"'
        f'{where_clause}'
    )


def resolve_sheet_query(sheet_name: str) -> str:
    """
    Chuyển đổi giá trị cột Sheet trong Options thành SQL:
    - Bắt đầu bằng SELECT → passthrough
    - Khớp pattern join rút gọn → gọi parse_join_expression
    - Còn lại → SELECT * FROM "<sheet_name>"
    """
    s = sheet_name.strip()
    if not s:
        return f'SELECT * FROM "{s}"'
    if s.lower().startswith("select"):
        return s
    if _JOIN_RE.match(s):
        return parse_join_expression(s)
    return f'SELECT * FROM "{validate_sql_identifier(s)}"'


def _parse_repeat_sheet_config(opt_config: dict) -> tuple[str, str, str]:
    """
    Phân tích cột Sheet dạng join rút gọn.
    Trả về (left_sheet, right_sheet, join_key).
    """
    sheet_expr = opt_config.get("sheet", "").strip()
    # Tách phần WHERE ra trước khi phân tích cấu trúc bảng/khóa join
    where_parts = re.split(r'\s+WHERE\s+', sheet_expr, flags=re.IGNORECASE)
    sheet_expr_clean = where_parts[0].strip()

    if "@" not in sheet_expr_clean:
        return validate_sql_identifier(sheet_expr_clean), "", ""

    join_part, key_raw = sheet_expr_clean.split("@", 1)
    join_part = join_part.strip()
    key_raw   = key_raw.strip()

    if "=" in key_raw:
        parts    = [validate_sql_identifier(k.strip()) for k in key_raw.split("=", 1)]
        join_key = f"{parts[0]} = {parts[1]}"
    else:
        join_key = validate_sql_identifier(key_raw)

    for sym, _ in _OP_MAP:
        if sym in join_part:
            left, right = join_part.split(sym.strip(), 1)
            return (
                validate_sql_identifier(left.strip()),
                validate_sql_identifier(right.strip()),
                join_key,
            )

    return validate_sql_identifier(sheet_expr_clean), "", ""
