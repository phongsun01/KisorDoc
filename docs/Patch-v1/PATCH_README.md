# Patch Notes – KisorDoc

## Cách áp dụng

Mỗi file patch là file Python hoàn chỉnh — **thay thế trực tiếp** file gốc cùng tên:

| File patch | Thay thế file gốc | Bug được fix |
|---|---|---|
| `patch_filters.py` | `filters.py` | filter_number logic ngược |
| `patch_dataset.py` | `dataset.py` | Mở Excel 2 lần mỗi sheet |
| `patch_table_copier.py` | `table_copier.py` | 5 bug (vMerge, hardcode, date format, hidden cols) |
| `patch_main.py` | Xem hướng dẫn bên dưới | Orphan code + DataSet reuse + merger error handling |

---

## Chi tiết từng fix

### `patch_filters.py` – filter_number
**Bug:** `str(value).replace(".", "").replace(",", ".")` bị sai với input US-style `"1,500,000"`.

**Fix:** Detect format trước (VN/US/mixed), rồi mới normalize.

```python
# TRƯỚC (lỗi với "1,500,000"):
num = float(re.sub(r"[^\d.,\-]", "", str(value).replace(".", "").replace(",", ".")))

# SAU:
# Detect dấu phân cách nghìn vs thập phân → normalize → parse
```

---

### `patch_dataset.py` – mở file 2 lần
**Bug:** `_load()` mở file để lấy `sheetnames`, sau đó `_sheet_to_dataframe()` mở lại lần 2.

**Fix:** Gộp thành 1 lần mở, truyền `worksheet` object trực tiếp vào `_ws_to_dataframe()`.

```python
# TRƯỚC: 2 lần open
wb = openpyxl.load_workbook(xlsx_file, read_only=True)          # lần 1
df = _sheet_to_dataframe(xlsx_file, stripped)                    # mở lại lần 2 bên trong

# SAU: 1 lần open
wb = openpyxl.load_workbook(xlsx_file, read_only=True, data_only=True)
df = _ws_to_dataframe(wb[sheet_name])                            # truyền ws object
```

---

### `patch_table_copier.py` – 5 bug

**Bug 1 – TABLE_PLACEHOLDER_RE hardcode:**
```python
# TRƯỚC:
TABLE_PLACEHOLDER_RE = re.compile(r"\{\{(DanhMucKoGia|DanhMuc)\}\}")

# SAU: nhận bất kỳ tên nào, lọc theo valid_keys từ tables_data
TABLE_PLACEHOLDER_RE = re.compile(r"\{\{(\w+)\}\}")
# valid_placeholder_keys được build từ sheet Tables
```

**Bug 2 – vMerge: ô phụ merge dọc bị skip:**
```python
# TRƯỚC: ô phụ bị bỏ qua hoàn toàn → Word bị thiếu ô → crash
if not is_master:
    continue

# SAU: tạo ô với vMerge=continue
if slave_type == "v":
    tc = _create_vmerge_continue_cell(NS)  # ← tạo ô đúng cách
    tr.append(tc)
```

**Bug 3 – vMerge condition sai:**
```python
# TRƯỚC: rs==1 vẫn tạo vMerge=continue → sai
if rs > 1:
    vMerge.set(val, "restart")
else:
    vMerge.set(val, "continue")  # ← sai

# SAU: chỉ set vMerge khi rs > 1
if rs > 1:
    vMerge.set(val, "restart")
# rs == 1: không có vMerge
```

**Bug 4 – Date format string:**
```python
# TRƯỚC: dùng Excel format làm strftime → sai
fmt = nf.replace("hh:mm:ss", "").strip()   # "DD/MM/YYYY"
return dt.strftime(fmt)                      # → "DD/07/YYYY" thay vì "01/07/2026"

# SAU: convert Excel format → Python strftime trước
py_fmt = _excel_date_fmt_to_strftime(nf)    # "DD/MM/YYYY" → "%d/%m/%Y"
return dt.strftime(py_fmt)                   # → "01/07/2026" ✓
```

**Bug 5 – _parse_hidden_cols chỉ nhận column letter:**
```python
# TRƯỚC: "DonGia,ThanhTien" → không ẩn được gì
openpyxl.utils.column_index_from_string("DonGia")  # → ValueError → bị bỏ qua

# SAU: thử column letter trước, nếu fail thì match với header_row dict
try:
    col_idx = openpyxl.utils.column_index_from_string(part)  # "A","B","C"
except Exception:
    # Thử match tên header (case-insensitive)
    for header_name, col_idx in header_row.items():
        if header_name.lower() == part.lower():
            cols.add(col_idx)
```

---

### `patch_main.py` – 3 fix

**Bug 1 – Orphan code:**
Xóa đoạn `if val is None: return ""...` nằm ngoài function (khoảng dòng 56–64).

**Bug 2 – DataSet tạo mới mỗi lần:**
```python
# TRƯỚC: chậm
async def run_batch(...):
    ds_local = DataSet(config)   # load lại toàn bộ Excel

# SAU: dùng global
async def run_batch(...):
    global config, ds            # ds đã được init() ở startup
```

**Bug 3 – merger.py thiếu error handling:**
Thêm `mail_merge_safe()` vào `merger.py`:
- Render vào file tạm trước
- Chỉ ghi đè file gốc khi thành công
- Trả về `(bool, error_str)` thay vì raise exception

---

## Thứ tự áp dụng

1. Backup toàn bộ code gốc
2. Copy `patch_filters.py` → `filters.py`
3. Copy `patch_dataset.py` → `dataset.py`
4. Copy `patch_table_copier.py` → `table_copier.py`
5. Sửa `main.py` theo hướng dẫn trong `patch_main.py`:
   - Xóa orphan code (~dòng 56-64)
   - Thay `run_batch()` bằng `run_batch_fixed()`
   - Thêm `mail_merge_safe()` vào `merger.py` (xem `MERGER_PATCH` trong patch_main.py)
6. Test với 1 file template đơn giản trước
7. Test với bảng có merged cells
