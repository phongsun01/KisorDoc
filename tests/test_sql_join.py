"""
tests/test_sql_join.py
────────────────────────
Unit test cho kisorlib/sql_join.py (Refactor v3.2 Phase B).

Bao phủ:
  - validate_sql_identifier
  - parse_join_expression (tất cả operators + edge cases)
  - resolve_sheet_query   (router chính — 0 test trước đây)
  - _parse_repeat_sheet_config
"""

import pytest

from kisorlib.sql_join import (
    _JOIN_RE,
    _OP_MAP,
    _parse_repeat_sheet_config,
    parse_join_expression,
    resolve_sheet_query,
    validate_sql_identifier,
)


# ══════════════════════════════════════════════
# validate_sql_identifier
# ══════════════════════════════════════════════

class TestValidateSqlIdentifier:
    def test_ascii_plain(self):
        assert validate_sql_identifier("GoiThau") == "GoiThau"

    def test_with_underscore(self):
        assert validate_sql_identifier("GoiThau_ID") == "GoiThau_ID"

    def test_vietnamese_accented(self):
        assert validate_sql_identifier("Tổ chuyên gia") == "Tổ chuyên gia"
        assert validate_sql_identifier("Đấu thầu") == "Đấu thầu"
        assert validate_sql_identifier("Nhà thầu") == "Nhà thầu"

    def test_space_allowed(self):
        assert validate_sql_identifier("Số hiệu gói thầu") == "Số hiệu gói thầu"

    def test_hyphen_dot_hash_allowed(self):
        assert validate_sql_identifier("Member-ID") == "Member-ID"
        assert validate_sql_identifier("Config.Date") == "Config.Date"
        assert validate_sql_identifier("#TempTable") == "#TempTable"

    def test_empty_string(self):
        assert validate_sql_identifier("") == ""

    def test_raises_on_single_quote(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            validate_sql_identifier("ID' OR '1'='1")

    def test_raises_on_double_quote(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            validate_sql_identifier('ID" OR "1"="1')

    def test_raises_on_semicolon(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            validate_sql_identifier("GoiThau; DROP TABLE Options; --")

    def test_raises_on_slash(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            validate_sql_identifier("Table/Name")

    def test_raises_on_backslash(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            validate_sql_identifier("Table\\Name")

    def test_raises_on_parenthesis(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            validate_sql_identifier("func()")


# ══════════════════════════════════════════════
# parse_join_expression
# ══════════════════════════════════════════════

class TestParseJoinExpression:

    # ── Operator coverage ──

    def test_left_join(self):
        sql = parse_join_expression("GoiThau <* TCGTTD @ GoiThau_ID")
        assert "LEFT JOIN" in sql
        assert '"GoiThau" LEFT JOIN "TCGTTD"' in sql
        assert '"GoiThau"."GoiThau_ID" = "TCGTTD"."GoiThau_ID"' in sql

    def test_right_join(self):
        sql = parse_join_expression("GoiThau *> TCGTTD @ GoiThau_ID")
        assert "RIGHT JOIN" in sql
        assert '"GoiThau" RIGHT JOIN "TCGTTD"' in sql

    def test_inner_join(self):
        sql = parse_join_expression("GoiThau * TCGTTD @ GoiThau_ID")
        assert "INNER JOIN" in sql
        assert '"GoiThau" INNER JOIN "TCGTTD"' in sql

    def test_full_outer_join(self):
        sql = parse_join_expression("GoiThau <*> TCGTTD @ GoiThau_ID")
        assert "FULL OUTER JOIN" in sql
        assert '"GoiThau" FULL OUTER JOIN "TCGTTD"' in sql

    # ── Operator priority: <*> phải được nhận trước <* ──

    def test_full_outer_not_confused_with_left(self):
        # <*> phải là FULL OUTER, không phải LEFT
        sql = parse_join_expression("A <*> B @ k")
        assert "FULL OUTER JOIN" in sql
        assert "LEFT JOIN" not in sql

    def test_left_not_confused_with_inner(self):
        # <* phải là LEFT, không phải INNER *
        sql = parse_join_expression("A <* B @ k")
        assert "LEFT JOIN" in sql
        assert "INNER JOIN" not in sql

    def test_right_not_confused_with_inner(self):
        # *> phải là RIGHT, không phải INNER *
        sql = parse_join_expression("A *> B @ k")
        assert "RIGHT JOIN" in sql
        assert "INNER JOIN" not in sql

    # ── Key khác tên (@ k1 = k2) ──

    def test_different_key_names_left(self):
        sql = parse_join_expression("GoiThau <* TCGTTD @ GoiThau_ID = MemberID")
        assert "LEFT JOIN" in sql
        assert '"GoiThau"."GoiThau_ID" = "TCGTTD"."MemberID"' in sql

    def test_different_key_names_full_outer(self):
        sql = parse_join_expression("Table1 <*> Table2 @ GoiThau_ID = MemberID")
        assert "FULL OUTER JOIN" in sql
        assert '"Table1"."GoiThau_ID" = "Table2"."MemberID"' in sql

    def test_different_key_names_right(self):
        sql = parse_join_expression("A *> B @ k1 = k2")
        assert "RIGHT JOIN" in sql
        assert '"A"."k1" = "B"."k2"' in sql

    def test_different_key_names_inner(self):
        sql = parse_join_expression("A * B @ k1 = k2")
        assert "INNER JOIN" in sql
        assert '"A"."k1" = "B"."k2"' in sql

    # ── Tên bảng tiếng Việt ──

    def test_vietnamese_table_names(self):
        sql = parse_join_expression("Gói thầu <* Thành viên @ ID")
        assert "LEFT JOIN" in sql
        assert '"Gói thầu"' in sql
        assert '"Thành viên"' in sql

    def test_vietnamese_right_table_only(self):
        sql = parse_join_expression("GoiThau <* Nhà thầu @ GoiThau_ID")
        assert "LEFT JOIN" in sql
        assert '"Nhà thầu"' in sql

    # ── SELECT passthrough ──

    def test_select_passthrough_uppercase(self):
        expr = "SELECT * FROM GoiThau"
        assert parse_join_expression(expr) == expr

    def test_select_passthrough_lowercase(self):
        expr = "select distinct name from T where x = 1"
        assert parse_join_expression(expr) == expr

    def test_select_passthrough_with_join(self):
        expr = "SELECT * FROM A LEFT JOIN B ON A.k = B.k"
        assert parse_join_expression(expr) == expr

    # ── Không có @ → SELECT đơn ──

    def test_no_at_sign_simple_table(self):
        assert parse_join_expression("GoiThau") == 'SELECT * FROM "GoiThau"'

    def test_no_at_sign_vietnamese(self):
        assert parse_join_expression("Đấu thầu") == 'SELECT * FROM "Đấu thầu"'

    # ── Extra whitespace ──

    def test_extra_spaces_around_operator(self):
        sql = parse_join_expression("  GoiThau  <*  TCGTTD  @  GoiThau_ID  ")
        assert "LEFT JOIN" in sql
        assert '"GoiThau"' in sql
        assert '"TCGTTD"' in sql
        assert '"GoiThau_ID"' in sql

    # ── Tên bảng có ký tự nguy hiểm → ValueError ──

    def test_raises_on_dangerous_left_table(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            parse_join_expression("Table'; DROP TABLE-- <* B @ k")

    def test_raises_on_dangerous_right_table(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            parse_join_expression("A <* B'; DROP-- @ k")

    def test_raises_on_dangerous_key(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            parse_join_expression("A <* B @ k'; DROP TABLE--")

    # ── SQL output structure ──

    def test_output_starts_with_select(self):
        sql = parse_join_expression("A <* B @ k")
        assert sql.strip().upper().startswith("SELECT")

    def test_output_has_on_clause(self):
        sql = parse_join_expression("A * B @ k")
        assert " ON " in sql

    def test_join_with_where_clause(self):
        sql = parse_join_expression("GoiThau * TCGTTD @ GoiThau_ID WHERE GoiThau.GoiThau_HTDT == 'DTRR'")
        assert 'SELECT * FROM "GoiThau" INNER JOIN "TCGTTD" ON "GoiThau"."GoiThau_ID" = "TCGTTD"."GoiThau_ID" WHERE GoiThau.GoiThau_HTDT = \'DTRR\'' in sql

    def test_simple_table_with_where_clause(self):
        sql = parse_join_expression("GoiThau WHERE GoiThau_HTDT == 'DTRR'")
        assert 'SELECT * FROM "GoiThau" WHERE GoiThau_HTDT = \'DTRR\'' in sql


# ══════════════════════════════════════════════
# resolve_sheet_query  (trước đây 0 test)
# ══════════════════════════════════════════════

class TestResolveSheetQuery:

    # ── SELECT passthrough ──

    def test_passthrough_select_uppercase(self):
        expr = "SELECT * FROM GoiThau WHERE ID = 1"
        assert resolve_sheet_query(expr) == expr

    def test_passthrough_select_lowercase(self):
        expr = "select distinct name from T"
        assert resolve_sheet_query(expr) == expr

    def test_passthrough_select_with_subquery(self):
        expr = "SELECT a, b FROM T1 JOIN T2 ON T1.k = T2.k"
        assert resolve_sheet_query(expr) == expr

    # ── Simple sheet name → SELECT * FROM ──

    def test_simple_ascii_sheet(self):
        assert resolve_sheet_query("GoiThau") == 'SELECT * FROM "GoiThau"'

    def test_simple_vietnamese_sheet(self):
        assert resolve_sheet_query("Đấu thầu") == 'SELECT * FROM "Đấu thầu"'

    def test_simple_sheet_with_space(self):
        assert resolve_sheet_query("Nhà thầu") == 'SELECT * FROM "Nhà thầu"'

    def test_simple_sheet_underscore(self):
        assert resolve_sheet_query("GoiThau_ID") == 'SELECT * FROM "GoiThau_ID"'

    # ── Whitespace trimming ──

    def test_leading_trailing_spaces(self):
        assert resolve_sheet_query("  GoiThau  ") == 'SELECT * FROM "GoiThau"'

    def test_tabs_stripped(self):
        assert resolve_sheet_query("\tGoiThau\t") == 'SELECT * FROM "GoiThau"'

    # ── Join expressions → parse_join_expression ──

    def test_left_join_routed(self):
        sql = resolve_sheet_query("GoiThau <* TCGTTD @ GoiThau_ID")
        assert "LEFT JOIN" in sql
        assert '"GoiThau"' in sql
        assert '"TCGTTD"' in sql

    def test_right_join_routed(self):
        sql = resolve_sheet_query("GoiThau *> TCGTTD @ GoiThau_ID")
        assert "RIGHT JOIN" in sql

    def test_inner_join_routed(self):
        sql = resolve_sheet_query("GoiThau * TCGTTD @ GoiThau_ID")
        assert "INNER JOIN" in sql

    def test_full_outer_join_routed(self):
        sql = resolve_sheet_query("GoiThau <*> TCGTTD @ GoiThau_ID")
        assert "FULL OUTER JOIN" in sql

    def test_join_with_different_keys(self):
        sql = resolve_sheet_query("GoiThau <* TCGTTD @ GoiThau_ID = MemberID")
        assert "LEFT JOIN" in sql
        assert '"GoiThau"."GoiThau_ID" = "TCGTTD"."MemberID"' in sql

    def test_join_vietnamese_tables(self):
        sql = resolve_sheet_query("Gói thầu <* Nhà thầu @ ID")
        assert "LEFT JOIN" in sql
        assert '"Gói thầu"' in sql
        assert '"Nhà thầu"' in sql

    # ── _JOIN_RE không match: không có @ hoặc không có space quanh operator ──

    def test_no_space_around_operator_not_join(self):
        # 'GoiThau<*TCGTTD@key' không match _JOIN_RE (yêu cầu \s+ quanh operator)
        # → rơi vào nhánh validate_sql_identifier → raise vì có ký tự *
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            resolve_sheet_query("GoiThau<*TCGTTD@key")

    def test_one_side_space_not_join(self):
        # 'GoiThau <*TCGTTD @ key' — thiếu space sau <* → không match _JOIN_RE
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            resolve_sheet_query("GoiThau <*TCGTTD @ key")

    # ── Dangerous input → ValueError ──

    def test_raises_on_injection_attempt(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            resolve_sheet_query("GoiThau'; DROP TABLE--")

    def test_raises_on_semicolon(self):
        with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
            resolve_sheet_query("GoiThau; DROP TABLE GoiThau; --")

    # ── Nhất quán với parse_join_expression ──

    def test_consistent_with_parse_join(self):
        """resolve_sheet_query khi match join phải cho kết quả = parse_join_expression."""
        expr = "GoiThau <*> TCGTTD @ GoiThau_ID"
        assert resolve_sheet_query(expr) == parse_join_expression(expr)

    def test_consistent_select_passthrough(self):
        expr = "SELECT * FROM GoiThau"
        assert resolve_sheet_query(expr) == parse_join_expression(expr)


# ══════════════════════════════════════════════
# _parse_repeat_sheet_config
# ══════════════════════════════════════════════

class TestParseRepeatSheetConfig:

    def test_left_join(self):
        assert _parse_repeat_sheet_config({"sheet": "GoiThau <* TCGTTD @ GoiThau_ID"}) == (
            "GoiThau", "TCGTTD", "GoiThau_ID"
        )

    def test_right_join(self):
        assert _parse_repeat_sheet_config({"sheet": "GoiThau *> TCGTTD @ GoiThau_ID"}) == (
            "GoiThau", "TCGTTD", "GoiThau_ID"
        )

    def test_inner_join(self):
        assert _parse_repeat_sheet_config({"sheet": "GoiThau * TCGTTD @ GoiThau_ID"}) == (
            "GoiThau", "TCGTTD", "GoiThau_ID"
        )

    def test_full_outer_join(self):
        assert _parse_repeat_sheet_config({"sheet": "GoiThau <*> TCGTTD @ GoiThau_ID"}) == (
            "GoiThau", "TCGTTD", "GoiThau_ID"
        )

    def test_different_key_names(self):
        left, right, key = _parse_repeat_sheet_config(
            {"sheet": "GoiThau <* TCGTTD @ GoiThau_ID = MemberID"}
        )
        assert left == "GoiThau"
        assert right == "TCGTTD"
        assert "GoiThau_ID" in key and "MemberID" in key

    def test_simple_sheet_no_join(self):
        left, right, key = _parse_repeat_sheet_config({"sheet": "GoiThau"})
        assert left == "GoiThau"
        assert right == ""
        assert key == ""

    def test_empty_sheet(self):
        assert _parse_repeat_sheet_config({"sheet": ""}) == ("", "", "")

    def test_missing_sheet_key(self):
        assert _parse_repeat_sheet_config({}) == ("", "", "")

    def test_vietnamese_table_names(self):
        left, right, key = _parse_repeat_sheet_config(
            {"sheet": "Gói thầu <* Nhà thầu @ GoiThau_ID"}
        )
        assert left == "Gói thầu"
        assert right == "Nhà thầu"
        assert key == "GoiThau_ID"

    def test_parse_repeat_sheet_with_where_clause(self):
        left, right, key = _parse_repeat_sheet_config(
            {"sheet": "GoiThau * TCGTTD @ GoiThau_ID WHERE GoiThau.GoiThau_HTDT == 'DTRR'"}
        )
        assert left == "GoiThau"
        assert right == "TCGTTD"
        assert key == "GoiThau_ID"

    def test_full_outer_not_confused_with_left(self):
        # <*> phải trả (left, right) đúng, không split lẫn với <*
        left, right, key = _parse_repeat_sheet_config(
            {"sheet": "A <*> B @ k"}
        )
        assert left == "A"
        assert right == "B"

    def test_symmetric_roundtrip_via_resolve(self):
        """left/right từ _parse_repeat_sheet_config phải xuất hiện trong SQL của resolve_sheet_query."""
        expr = "GoiThau <* TCGTTD @ GoiThau_ID"
        left, right, _ = _parse_repeat_sheet_config({"sheet": expr})
        sql = resolve_sheet_query(expr)
        assert f'"{left}"' in sql
        assert f'"{right}"' in sql


# ══════════════════════════════════════════════
# _JOIN_RE
# ══════════════════════════════════════════════

class TestJoinRe:
    """Kiểm tra regex guard trước khi route vào parse_join_expression."""

    def test_matches_left(self):
        assert _JOIN_RE.match("GoiThau <* TCGTTD @ GoiThau_ID")

    def test_matches_right(self):
        assert _JOIN_RE.match("A *> B @ k")

    def test_matches_inner(self):
        assert _JOIN_RE.match("A * B @ k")

    def test_matches_full_outer(self):
        assert _JOIN_RE.match("A <*> B @ k")

    def test_no_match_simple_name(self):
        assert not _JOIN_RE.match("GoiThau")

    def test_no_match_select(self):
        assert not _JOIN_RE.match("SELECT * FROM GoiThau")

    def test_no_match_no_at(self):
        assert not _JOIN_RE.match("A <* B")

    def test_no_match_no_space_around_op(self):
        # Yêu cầu \s+ quanh operator
        assert not _JOIN_RE.match("A<*B@k")
        assert not _JOIN_RE.match("GoiThau<*TCGTTD@key")
