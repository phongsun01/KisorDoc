"""
kisorlib/app_helpers.py
───────────────────────
Các hàm helper dùng chung giữa app.py và engine.py.
Tách ra để tránh circular import (engine không import app.py).

Nội dung được cut từ app.py:
  - NestedVal      (dòng 577–589)
  - make_nested_dict (dòng 592–629)

Sau khi tạo file này:
  - app.py: thay 2 định nghĩa trên bằng:
        from kisorlib.app_helpers import NestedVal, make_nested_dict
  - engine.py: dùng:
        from .app_helpers import make_nested_dict
"""


class NestedVal(dict):
    """
    Dict con có thể dùng như string — dùng để hỗ trợ cú pháp
    Jinja2 dạng {{ Bien.SubKey }} trong docxtpl.
    """
    def __init__(self, val, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self["_val"] = val

    def __str__(self):
        return str(self.get("_val", ""))

    def __repr__(self):
        return self.__str__()

    def __html__(self):
        return self.__str__()


def make_nested_dict(flat_dict: dict) -> dict:
    """
    Chuyển dict phẳng có key dạng "A.B.C" thành dict lồng nhau,
    hỗ trợ cú pháp {{ A.B.C }} trong Jinja2/docxtpl.

    Ví dụ:
        {"TenGoiThau": "ABC", "NguoiKy.HoTen": "Nguyen Van A"}
        →
        {"TenGoiThau": "ABC", "NguoiKy": NestedVal("") {"HoTen": "Nguyen Van A"}}
    """
    nested = {}

    for key, value in flat_dict.items():
        parts = key.split(".")
        d = nested
        for part in parts[:-1]:
            if part not in d:
                d[part] = {}
            elif not isinstance(d[part], dict):
                d[part] = NestedVal(d[part])
            d = d[part]

        last_part = parts[-1]
        if last_part in d:
            if isinstance(d[last_part], dict):
                d[last_part]["_val"] = value
            else:
                d[last_part] = value
        else:
            d[last_part] = value

    # Post-process: chuyển dict thường có "_val" thành NestedVal
    def _convert(obj):
        if isinstance(obj, dict):
            for k, v in list(obj.items()):
                obj[k] = _convert(v)
            if "_val" in obj and not isinstance(obj, NestedVal):
                nv = NestedVal(obj["_val"])
                for k, v in obj.items():
                    if k != "_val":
                        nv[k] = v
                return nv
        return obj

    for k, v in list(nested.items()):
        nested[k] = _convert(v)

    return nested
