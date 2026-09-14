# MGV — Phần mềm đề xuất đặt phụ tùng dịch vụ

So sánh tồn kho hiện tại với tiêu chuẩn Min/Max của HN & HCM, đề xuất danh sách
linh kiện cần đặt và xuất ra Excel đúng format công ty.

## Cài đặt
```bash
pip install -r requirements.txt
```

## Chạy
```bash
streamlit run app.py
```
Trình duyệt mở tại http://localhost:8501

## Cách dùng
1. (Tuỳ chọn) Cập nhật tiêu chuẩn ở sidebar, hoặc cứ dùng `standard.xlsx` đi kèm.
2. Import file tồn kho (sheet `PBI_Export`, xuất từ Power BI).
3. Chọn khu vực: **HN / HCM / Cả nước**.
4. Bấm **Phân tích** → xem/chỉnh bảng đề xuất, thêm item thủ công nếu cần.
5. Bấm **Tạo file đặt hàng** → tải Excel.

## Quy tắc tính (khớp tiêu chuẩn đã thống nhất)
- **HN** = tổng tồn các kho mã đuôi `23`; **HCM** = đuôi `21`.
- Đuôi khác (`10`, `30`, `31`…) = hàng đi đường/trung chuyển → **không tính** vào tồn.
- Khớp mã 2 tầng: **Part Number 2026** → **Former P/N** (fallback).
- Trigger: `Onhand ≤ Min` → đặt `Max − Onhand`, làm tròn LÊN đơn vị nguyên.
- **Onhand HN** và **Onhand HCM** luôn điền ở mọi dòng để đối chiếu.
- Sort: **Stago trước, Sebia sau**; trong brand **A–Z theo Item Name**.
- **Cost** = `Quantity × Unit Price (USD)`, cột nằm ngay sau **Quantity**.
- **Nguồn giá**: DUY NHẤT từ `partlist.xlsx` (cột `Unit Price`), tra theo mã 2026 và Former P/N. `standard.xlsx` KHÔNG còn cột giá. Item nhập tay ngoài standard vẫn có giá nếu partlist có.
- Item nhập tay: ô **Item Name** để trống sẽ **tự tra tên** từ `partlist.xlsx` (SPP master); tên nhập tay được ưu tiên nếu có.

- **Cả nước** → xuất **2 sheet độc lập** `HN` và `HCM`.

## Cập nhật tiêu chuẩn
Sửa/thay file `standard.xlsx` (giữ đúng cấu trúc cột như `All_Standard_Inventory`):
`Part Numer (2026) | Part name | Former P/N | Analyzer | Brand | HAN Min | HAN Max | HCM Min | HCM Max`
> Engine đọc cột **theo tên header** nên có thể thêm/bớt cột mà không vỡ. Giá KHÔNG còn ở standard — sửa giá tại `partlist.xlsx`.
Không cần sửa code.

## Cấu trúc mã nguồn
- `engine.py`  — logic thuần (nạp tiêu chuẩn, gom tồn, khớp mã, tính đặt hàng). Test được độc lập.
- `writer.py`  — xuất Excel chuẩn MGV (14 cột, dropdown Item Group).
- `app.py`     — giao diện Streamlit.
- `standard.xlsx` — tiêu chuẩn tồn kho (có cột `Unit Price (USD)`).
- `partlist.xlsx` — SPP master: tra **tên** và **đơn giá** từ mã (cột `Unit Price`). Cập nhật giá tại đây.


## Đăng nhập (mật khẩu qua Streamlit secrets)
App có màn đăng nhập ở đầu. Mật khẩu KHÔNG nằm trong code mà đọc từ Streamlit **secrets**.

Chạy LOCAL: tạo file `.streamlit/secrets.toml` (xem `secrets.toml.example`):
```toml
app_password = "mat_khau_cua_ban"
```
Hoặc nhiều tài khoản:
```toml
[passwords]
nhan = "..."
tuan = "..."
```
Trên Streamlit Cloud: mở app -> **Settings -> Secrets** -> dán nội dung trên -> Save (app tự khởi động lại).
File `secrets.toml` thật đã được `.gitignore` loại trừ, không bị đẩy lên GitHub.

## Cập nhật giá
Sửa cột `Unit Price` trong `partlist.xlsx` (mỗi mã một dòng). Đây là nguồn giá duy nhất; `standard.xlsx` không còn cột giá.
