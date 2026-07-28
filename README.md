# KisorDoc-AI

Công cụ Python xử lý hàng loạt tài liệu Word (Mail Merge & Copy bảng từ Excel sang Word), được xây dựng để thay thế hoàn toàn bot UiPath cũ.

## Tính năng chính

1. **Mail Trộn thư (Mail Merge):**
   - Hỗ trợ trộn dữ liệu từ Excel vào Word qua template định dạng Jinja2 `{{ TenBien }}`.
   - Hỗ trợ các bộ lọc định dạng: `|date` (dd/MM/yyyy), `|date_long` (ngày... tháng... năm...), `|number` (định dạng số phân tách hàng nghìn), `|upper` (chữ in hoa).
   - Tự động phân tách và xử lý các biến văn bản độc lập với biến ngày tháng cùng tiền tố (Ví dụ: `KHLCNT_TTr` và `KHLCNT_TTr_Date`).
   - Tự động phân tích các biến lồng nhau dạng dấu chấm (Ví dụ: `{{KHLCNT_TTr.Dvi}}`, `{{DuToan.NguoiLap}}`).

2. **Copy bảng Excel sang Word:**
   - Sao chép một vùng dữ liệu (Range) từ Excel và chèn vào vị trí placeholder `{{DanhMuc}}` hoặc `{{DanhMucKoGia}}` dưới dạng bảng Word thật.
   - Giữ nguyên định dạng gốc: gộp ô (merged cells), màu nền, viền bảng, chiều cao/chiều rộng và căn lề.
   - Tự động chuẩn hóa từ các ký hiệu ngoặc đơn `{}` hoặc không ngoặc trong Excel.

3. **Giao diện Web Local (Gradio):**
   - Thao tác trực quan qua 3 bước: Chọn Gói thầu -> Chọn template -> Chạy & Xem log.

## Cấu trúc thư mục

```text
{ProjectPath}/
├── 1. Data/            # Chứa các file dữ liệu Excel (.xlsx)
├── 2. Templates/       # Chứa các template Word (.docx), chia theo Opt1/Opt2
├── 3. Files/           # Thư mục đầu ra (Output)
└── Config-5.txt        # Cấu hình gốc (đọc từ %LOCALAPPDATA%\UiPathProjectConfigs\)
```

## Hướng dẫn cài đặt và khởi chạy

1. Cài đặt các thư viện phụ thuộc:
   ```bash
   pip install -r requirements.txt
   ```
2. Cấu hình các biến môi trường trong file `.env` tại thư mục gốc để ghi đè `Config-5.txt` nếu cần.
3. Chạy ứng dụng:
   ```bash
   python main.py
   ```
   Ứng dụng sẽ khởi chạy tại cổng mặc định `http://127.0.0.1:7864`.
