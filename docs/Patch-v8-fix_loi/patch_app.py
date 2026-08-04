"""
PATCH app.py — 3 thay đổi, áp dụng theo thứ tự
=================================================

THAY ĐỔI 1: Thêm import vào cuối block import hiện tại (sau dòng "from docxtpl import DocxTemplate")
─────────────────────────────────────────────────────
THÊM dòng này:

    from kisorlib.app_helpers import NestedVal, make_nested_dict


THAY ĐỔI 2: XÓA class NestedVal (dòng 577–589 trong app.py)
─────────────────────────────────────────────────────
XÓA toàn bộ đoạn này:

    class NestedVal(dict):
        def __init__(self, val, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self["_val"] = val

        def __str__(self):
            return str(self.get("_val", ""))

        def __repr__(self):
            return self.__str__()

        def __html__(self):
            return self.__str__()


THAY ĐỔI 3: XÓA hàm make_nested_dict (dòng 592–629 trong app.py)
─────────────────────────────────────────────────────
XÓA toàn bộ đoạn này:

    def make_nested_dict(flat_dict: dict) -> dict:
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

        # Post-process: Convert any standard dict that has "_val" key into NestedVal
        def convert_to_nested_val(obj):
            if isinstance(obj, dict):
                for k, v in list(obj.items()):
                    obj[k] = convert_to_nested_val(v)
                if "_val" in obj and not isinstance(obj, NestedVal):
                    nv = NestedVal(obj["_val"])
                    for k, v in obj.items():
                        if k != "_val":
                            nv[k] = v
                    return nv
            return obj

        for k, v in list(nested.items()):
            nested[k] = convert_to_nested_val(v)

        return nested


=================================================
Kết quả sau khi áp dụng: app.py ngắn hơn ~55 dòng,
NestedVal + make_nested_dict nằm trong app_helpers.py,
cả app.py lẫn engine.py đều import từ đó.
=================================================
"""
