import math
import re
import ast

_JOIN_RE = re.compile(r'.+\s+(?:<\*>|<\*|\*>|\*)\s+.+\s*@\s*.+', re.DOTALL)

# Bảng ánh xạ toán tử join rút gọn → SQL JOIN (thứ tự ưu tiên: <*> trước <* trước *> trước *)
_OP_MAP = [
    (" <*>", "FULL OUTER JOIN"),
    (" <*",  "LEFT JOIN"),
    (" *>",  "RIGHT JOIN"),
    (" *",   "INNER JOIN"),
]


def _str(val, default="") -> str:
    if val is None:
        return default
    if isinstance(val, float):
        if math.isnan(val):
            return default
    return str(val).strip()


def validate_sql_identifier(name: str) -> str:
    if not name:
        return ""
    # Whitelist pattern for secure column/table name formatting
    pattern = r"^[A-Za-z0-9_\s\-\.\#\u00C0-\u1EF9]+$"
    if not re.match(pattern, name):
        raise ValueError(f"⚠️ Cảnh báo bảo mật: Phát hiện ký tự không hợp lệ trong tên bảng hoặc cột (SQL Injection Risk): '{name}'")
    return name



def clean_config_key(key: str) -> str:
    clean = key.strip("<>{}| ")
    
    # Maps dot date suffixes to _Date suffix
    for suffix in (".Date.Long", ".Date.long", ".date_long"):
        if clean.endswith(suffix):
            return clean[:-len(suffix)] + "_Date"
            
    if clean.endswith(".Date") or clean.endswith(".date"):
        return clean[:-5] + "_Date"
        
    if clean.endswith(".Day") or clean.endswith(".day"):
        return clean[:-4] + "_Date"
        
    if clean.endswith(".Month") or clean.endswith(".month"):
        return clean[:-6] + "_Date"
        
    if clean.endswith(".Year") or clean.endswith(".year"):
        return clean[:-5] + "_Date"
        
    # Other standard modifiers
    for suffix in (".Upper", ".upper", ".Number", ".number"):
        if clean.endswith(suffix):
            clean = clean[:-len(suffix)]
            break
            
    if "|" in clean:
        clean = clean.split("|")[0].strip()
        
    return clean


def _parse_row_range(s: str) -> tuple[int, int] | None:
    if not s or not s.strip():
        return None
    m = re.match(r"^(\d+)-(\d+)$", s.strip())
    if m:
        try:
            start = int(m.group(1))
            end = int(m.group(2))
            if start <= end:
                return start, end
        except ValueError:
            pass
    return None


def safe_format(pattern: str, row: dict) -> str:
    if not pattern:
        return ""
    res = pattern
    placeholders = re.findall(r"\{(.*?)\}", pattern)
    for p in placeholders:
        val = _str(row.get(p, ""))
        res = res.replace(f"{{{p}}}", val)
    return res


def _safe_eval_condition(expr: str, context: dict) -> bool:
    """
    AST-giới hạn: chỉ cho phép so sánh, phép toán số học, boolean, biến và hằng.
    Không cho phép call/attribute/import → không thể inject.
    """
    node = ast.parse(expr, mode="eval")

    def eval_node(n):
        if isinstance(n, ast.Constant):
            return n.value
        if isinstance(n, ast.Tuple) or isinstance(n, ast.List) or isinstance(n, ast.Set):
            return tuple(eval_node(e) for e in n.elts)
        if isinstance(n, ast.Name):
            if n.id not in context:
                raise ValueError(f"Biến '{n.id}' không tồn tại")
            return context[n.id]
        if isinstance(n, ast.BinOp):
            left = eval_node(n.left)
            right = eval_node(n.right)
            if isinstance(n.op, ast.Add):
                return left + right
            if isinstance(n.op, ast.Sub):
                return left - right
            if isinstance(n.op, ast.Mult):
                return left * right
            if isinstance(n.op, ast.Div):
                return left / right
            if isinstance(n.op, ast.FloorDiv):
                return left // right
            if isinstance(n.op, ast.Mod):
                return left % right
            if isinstance(n.op, ast.Pow):
                return left ** right
            raise ValueError("Toán tử số học không được hỗ trợ")
        if isinstance(n, ast.UnaryOp):
            operand = eval_node(n.operand)
            if isinstance(n.op, ast.Not):
                return not operand
            if isinstance(n.op, ast.USub):
                return -operand
            if isinstance(n.op, ast.UAdd):
                return +operand
            raise ValueError("Toán tử một ngôi không được hỗ trợ")
        if isinstance(n, ast.BoolOp):
            if isinstance(n.op, ast.And):
                return all(eval_node(v) for v in n.values)
            if isinstance(n.op, ast.Or):
                return any(eval_node(v) for v in n.values)
            raise ValueError("Toán tử boolean không được hỗ trợ")
        if isinstance(n, ast.Compare):
            left = eval_node(n.left)
            for op, comparator in zip(n.ops, n.comparators):
                right = eval_node(comparator)
                if isinstance(op, ast.Eq):
                    res = left == right
                elif isinstance(op, ast.NotEq):
                    res = left != right
                elif isinstance(op, ast.Lt):
                    res = left < right
                elif isinstance(op, ast.LtE):
                    res = left <= right
                elif isinstance(op, ast.Gt):
                    res = left > right
                elif isinstance(op, ast.GtE):
                    res = left >= right
                elif isinstance(op, ast.In):
                    res = left in right
                elif isinstance(op, ast.NotIn):
                    res = left not in right
                else:
                    raise ValueError("Toán tử so sánh không được hỗ trợ")
                if not res:
                    return False
                left = right
            return True
        if isinstance(n, ast.IfExp):
            return eval_node(n.body) if eval_node(n.test) else eval_node(n.orelse)
        raise ValueError(f"Biểu thức không được phép: {type(n).__name__}")

    return bool(eval_node(node.body))


def parse_join_expression(expr: str) -> str:
    """
    Cú pháp rút gọn cho cột Sheet trong Options:
      Table1 <* Table2 @ key           → LEFT JOIN, cùng tên cột
      Table1 <* Table2 @ key1 = key2   → LEFT JOIN, khác tên cột
      Table1 *> Table2 @ key           → RIGHT JOIN
      Table1 * Table2 @ key            → INNER JOIN
      Table1 <*> Table2 @ key          → FULL OUTER JOIN
      SELECT ...                        → truyền thẳng cho DuckDB
    """
    s = expr.strip()
    if s.lower().startswith("select"):
        return s
    if "@" not in s:
        return f'SELECT * FROM "{validate_sql_identifier(s)}"'

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
        return f'SELECT * FROM "{validate_sql_identifier(s)}"'

    if "=" in key_raw:
        k1, k2 = [validate_sql_identifier(k.strip()) for k in key_raw.split("=", 1)]
    else:
        k1 = k2 = validate_sql_identifier(key_raw)

    return (
        f'SELECT * FROM "{t1}" {join_type} "{t2}" '
        f'ON "{t1}"."{k1}" = "{t2}"."{k2}"'
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
    Phân tích cột Sheet dạng join rút gọn:
      GoiThau * TCGTTD @ GoiThau_ID          → INNER JOIN
      GoiThau <* TCGTTD @ GoiThau_ID         → LEFT JOIN
      GoiThau *> TCGTTD @ GoiThau_ID         → RIGHT JOIN
      GoiThau <*> TCGTTD @ GoiThau_ID        → FULL OUTER JOIN
    Trả về (left_sheet, right_sheet, join_key).
    Dùng cùng _OP_MAP với parse_join_expression để tránh split sai khi gặp <*> hay *>.
    """
    sheet_expr = opt_config.get("sheet", "").strip()

    # Không có dấu @ → không phải join → trả về sheet đơn
    if "@" not in sheet_expr:
        return validate_sql_identifier(sheet_expr), "", ""

    join_part, key_raw = sheet_expr.split("@", 1)
    join_part = join_part.strip()
    
    key_raw = key_raw.strip()
    if "=" in key_raw:
        parts = [validate_sql_identifier(k.strip()) for k in key_raw.split("=", 1)]
        join_key = f"{parts[0]} = {parts[1]}"
    else:
        join_key = validate_sql_identifier(key_raw)

    # Dùng chung constant _OP_MAP (theo thứ tự ưu tiên: <*> trước <* trước *> trước *)
    for sym, _ in _OP_MAP:
        if sym in join_part:
            left, right = join_part.split(sym.strip(), 1)
            return validate_sql_identifier(left.strip()), validate_sql_identifier(right.strip()), join_key

    # Không khớp operator nào → trả về sheet đơn
    return validate_sql_identifier(sheet_expr), "", ""


def _parse_repeat_key_id(key_id_expr: str) -> tuple[str, str]:
    """
    Phân tích cấu hình KeyId dạng: 'GoiThau_ID | CCCD'
    Trả về (left_key, right_key).
    """
    if not key_id_expr:
        return "ID", "ID"
    if "|" in key_id_expr:
        parts = key_id_expr.split("|", 1)
        return validate_sql_identifier(parts[0].strip()), validate_sql_identifier(parts[1].strip())
    val = validate_sql_identifier(key_id_expr.strip())
    return val, val


def _parse_price(val) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if math.isnan(val):
            return None
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None
