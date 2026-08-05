import math
import pytest
from kisorlib.utils import (
    _parse_price,
    clean_config_key,
    parse_join_expression,
    _parse_repeat_key_id,
    _parse_repeat_sheet_config,
    _parse_row_range,
    _str,
    safe_format,
    _safe_eval_condition,
    validate_sql_identifier,
)


def test_parse_price():
    assert _parse_price(None) is None
    assert _parse_price(float("nan")) is None
    assert _parse_price("1.500.000") == 1500000.0
    assert _parse_price("") is None
    assert _parse_price(1234.56) == 1234.56
    assert _parse_price("  ") is None
    assert _parse_price("1.234,56") == 1234.56  # test comma replacement


def test_clean_config_key():
    assert clean_config_key("Key.Date.Long") == "Key_Date"
    assert clean_config_key("Key.Date") == "Key_Date"
    assert clean_config_key("Key.upper") == "Key"
    assert clean_config_key("Key.Upper") == "Key"
    assert clean_config_key("Key.Number") == "Key"
    assert clean_config_key("Key | SomeOtherPart") == "Key"
    assert clean_config_key("<Key>") == "Key"
    assert clean_config_key("{Key}") == "Key"


def test_parse_join_expression():
    # <*> FULL OUTER JOIN
    expr_full = "Table1 <*> Table2 @ key"
    assert "FULL OUTER JOIN" in parse_join_expression(expr_full)
    assert '"Table1" FULL OUTER JOIN "Table2"' in parse_join_expression(expr_full)
    assert '"Table1"."key" = "Table2"."key"' in parse_join_expression(expr_full)

    # <* LEFT JOIN
    expr_left = "Table1 <* Table2 @ key1 = key2"
    assert "LEFT JOIN" in parse_join_expression(expr_left)
    assert '"Table1"."key1" = "Table2"."key2"' in parse_join_expression(expr_left)

    # *> RIGHT JOIN
    expr_right = "Table1 *> Table2 @ key"
    assert "RIGHT JOIN" in parse_join_expression(expr_right)

    # * INNER JOIN
    expr_inner = "Table1 * Table2 @ key"
    assert "INNER JOIN" in parse_join_expression(expr_inner)

    # SELECT passthrough
    expr_sql = "SELECT * FROM Table1"
    assert parse_join_expression(expr_sql) == expr_sql


def test_parse_join_expression_edge_cases():
    # <*> FULL OUTER JOIN với key khác tên (@ k1 = k2)
    sql = parse_join_expression("Table1 <*> Table2 @ GoiThau_ID = MemberID")
    assert "FULL OUTER JOIN" in sql
    assert '"Table1"."GoiThau_ID" = "Table2"."MemberID"' in sql

    # *> RIGHT JOIN với key khác tên (@ k1 = k2)
    sql = parse_join_expression("Table1 *> Table2 @ GoiThau_ID = MemberID")
    assert "RIGHT JOIN" in sql
    assert '"Table1"."GoiThau_ID" = "Table2"."MemberID"' in sql

    # * INNER JOIN với key khác tên
    sql = parse_join_expression("Table1 * Table2 @ GoiThau_ID = MemberID")
    assert "INNER JOIN" in sql
    assert '"Table1"."GoiThau_ID" = "Table2"."MemberID"' in sql

    # Không có @ → trả về SELECT đơn
    assert parse_join_expression("Table1") == 'SELECT * FROM "Table1"'


def test_parse_repeat_sheet_config():
    # Giá trị Options.Sheet dạng join rút gọn → (left_sheet, right_sheet, join_key)
    assert _parse_repeat_sheet_config({"sheet": "GoiThau <* TCGTTD @ GoiThau_ID"}) == ("GoiThau", "TCGTTD", "GoiThau_ID")
    assert _parse_repeat_sheet_config({"sheet": "GoiThau <*> TCGTTD @ GoiThau_ID"}) == ("GoiThau", "TCGTTD", "GoiThau_ID")
    assert _parse_repeat_sheet_config({"sheet": "GoiThau *> TCGTTD @ GoiThau_ID"}) == ("GoiThau", "TCGTTD", "GoiThau_ID")
    assert _parse_repeat_sheet_config({"sheet": "GoiThau * TCGTTD @ GoiThau_ID"}) == ("GoiThau", "TCGTTD", "GoiThau_ID")
    # Khác tên cột khóa
    assert _parse_repeat_sheet_config({"sheet": "GoiThau <* TCGTTD @ GoiThau_ID = MemberID"}) == ("GoiThau", "TCGTTD", "GoiThau_ID = MemberID")
    # Không phải join → sheet đơn, right rỗng
    assert _parse_repeat_sheet_config({"sheet": "GoiThau"}) == ("GoiThau", "", "")
    assert _parse_repeat_sheet_config({}) == ("", "", "")


def test_parse_repeat_key_id():
    assert _parse_repeat_key_id("") == ("ID", "ID")
    assert _parse_repeat_key_id("GoiThau_ID") == ("GoiThau_ID", "GoiThau_ID")
    assert _parse_repeat_key_id("GoiThau_ID | CCCD") == ("GoiThau_ID", "CCCD")
    assert _parse_repeat_key_id("  GoiThau_ID  |  CCCD  ") == ("GoiThau_ID", "CCCD")


def test_parse_row_range():
    assert _parse_row_range("2-97") == (2, 97)
    assert _parse_row_range("10-5") is None
    assert _parse_row_range("abc") is None
    assert _parse_row_range("") is None


def test_safe_eval_condition():
    assert _safe_eval_condition("var_0 == 'Opt1'", {"var_0": "Opt1"}) is True
    assert _safe_eval_condition("var_0 > 1000000", {"var_0": 1500000}) is True
    assert _safe_eval_condition("var_0 > 1000000 and var_1 == 'Active'", {"var_0": 1500000, "var_1": "Active"}) is True
    assert _safe_eval_condition("var_0 > 1000000 and var_1 == 'Active'", {"var_0": 1500000, "var_1": "Inactive"}) is False
    assert _safe_eval_condition("var_0 > 1000000 or var_1 == 'Active'", {"var_0": 500000, "var_1": "Active"}) is True
    assert _safe_eval_condition("var_0 > 1000000 or var_1 == 'Active'", {"var_0": 500000, "var_1": "Inactive"}) is False
    with pytest.raises(ValueError):
        _safe_eval_condition("var_x == 1", {})


def test_validate_sql_identifier():
    # Đúng chuẩn whitelist
    assert validate_sql_identifier("GoiThau") == "GoiThau"
    assert validate_sql_identifier("GoiThau_ID") == "GoiThau_ID"
    assert validate_sql_identifier("Tổ chuyên gia") == "Tổ chuyên gia"  # Tiếng Việt hợp lệ
    assert validate_sql_identifier("Số hiệu gói thầu") == "Số hiệu gói thầu"  # Khoảng trắng
    assert validate_sql_identifier("Member-ID") == "Member-ID"  # Dấu gạch ngang
    assert validate_sql_identifier("Config.Date") == "Config.Date"  # Dấu chấm
    assert validate_sql_identifier("#TempTable") == "#TempTable"  # Dấu thăng
    assert validate_sql_identifier("") == ""

    # Có ký tự nguy hiểm (SQL Injection)
    with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
        validate_sql_identifier("GoiThau; DROP TABLE Options; --")
    with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
        validate_sql_identifier("ID' OR '1'='1")
    with pytest.raises(ValueError, match="Cảnh báo bảo mật"):
        validate_sql_identifier("ID\" OR \"1\"=\"1")

