# BS Order Generator

Web tool local: đọc đơn hàng từ **Google Sheet** (hoặc file xlsx/csv upload) và tự tạo file
import đơn hàng chuẩn template BigSeller (`BS-order-import-template-260728.xlsx`).

## Chạy tool

```bash
./run.sh
```

Mở http://localhost:8787 trong trình duyệt.

## Cách dùng

1. **Nguồn dữ liệu** — chọn *Google Sheets* (cần đăng nhập Google, xem bên dưới) hoặc
   *Upload file* (.xlsx/.csv export từ Google Sheet).
2. **Chọn sheet** — các sheet trùng tên trong danh sách "Bỏ qua" (mặc định `Báo giá`)
   và các sheet **không có cột `BS - SKU`** sẽ tự được bỏ chọn. Chỉ những sheet có
   cột `BS - SKU` / `BS - Số lượng` mới được chọn sẵn. Mỗi sheet = 1 đơn hàng.
3. **Xem trước** — tool tự:
   - Sinh **ID đơn hàng** theo tên sheet. Tiền tố phân loại `Báo giá` / `Đơn hàng`
     (và dấu `-`, `()`) được loại trước khi đặt tên, vì đó không phải tên khách:
     - tên có `Aqua` → `Aqua_<tên>_<yyyymmdd>_<hhmm>`
       (vd `Báo giá - Lukibum Aqua` → `Aqua_Lukibum_20261007_1452`)
     - tên có `CTV` → `CTV_<tên>_<yyyymmdd>_<hhmm>`
     - còn lại → `KH_<tên>_<yyyymmdd>_<hhmm>`
       (vd `Đơn hàng (Bum)` → `KH_Bum_20261007_1452`)
   - Lấy **Tên người nhận** từ tên sheet (loại tiền tố phân loại + chữ Aqua/CTV)
   - Đọc cột `BS - SKU` và `BS - Số lượng`
   - Dòng thiếu SKU / số lượng ≤ 0 bị đánh đỏ và bỏ qua
4. **Xuất file** — **1 file xlsx duy nhất** (`BS-orders_yyyymmdd_hhmm.xlsx`) chứa tất cả đơn,
   các dòng của từng đơn được fill liên tiếp nhau. File được ghi trực tiếp vào template
   BigSeller nên giữ nguyên định dạng app trung gian yêu cầu.
   Cột được điền: `ID Đơn hàng`, `Tên người nhận`, `SKU gian hàng`, `Đơn giá = 0`, `Số lượng`.

## Đổi template

Template nằm trong thư mục **`template/`** — bỏ file `.xlsx` template mới vào đó là xong,
không cần sửa code. Tool tự chọn file `.xlsx` mới nhất trong folder và **tìm cột theo tên
header** (`ID Đơn hàng`, `Tên người nhận`, `SKU gian hàng`, `Đơn giá`, `Số lượng`), nên
template mới chỉ cần giữ đúng các tên cột này ở hàng đầu (trong 5 hàng đầu tiên).

## Cấu hình Google Sheets (OAuth)

Tool chỉ dùng quyền **đọc** (`spreadsheets.readonly`), không cần quyền Drive.

1. Vào [Google Cloud Console](https://console.cloud.google.com/) → tạo (hoặc chọn) 1 project.
2. *APIs & Services → OAuth consent screen*: chọn **External**, điền tên app, email bạn.
   Scope để mặc định (tool tự khai báo `spreadsheets.readonly` khi đăng nhập).
3. *APIs & Services → Credentials → Create credentials → OAuth client ID*:
   - Application type: **Web application**
   - Authorized JavaScript origins: `http://localhost:8787`
   - Authorized redirect URIs: `http://localhost:8787/oauth/callback`
4. Tải file JSON về, đặt vào thư mục tool với tên **`credentials.json`**.
5. Mở tool, bấm **Đăng nhập Google**, cấp quyền đọc. Token lưu lại trong `tokens.json`,
   các lần sau không cần đăng nhập lại.

> Nếu chưa kịp cấu hình OAuth: dùng tab **Upload file** — export Google Sheet ra
> `.xlsx` (File → Download) rồi tải lên, mọi tính năng khác vẫn hoạt động.

## Bản web public (Cloudflare Pages)

Thư mục **`site/`** là bản static — mọi logic chạy trong trình duyệt, không cần server:

- **Google Sheets**: dán link sheet đã bật chia sẻ **"Ai có link đều xem được"**
  (không cần OAuth, không cần đăng nhập Google). Web tải cả workbook về bằng endpoint
  export công khai của Google rồi parse bằng SheetJS, tự chọn sẵn các sheet có cột
  `BS - SKU`.
- **Upload file**: .xlsx/.csv export từ Google Sheet
- **Template**: mặc định là `site/template.xlsx` (bỏ file template mới vào đó rồi
  deploy lại); hoặc upload template tùy chỉnh ngay trên web (lưu trong trình duyệt)

### Deploy lên Cloudflare Pages

1. Đẩy repo lên GitHub/GitLab.
2. Cloudflare Dashboard → **Workers & Pages → Create → Pages → Connect to Git**.
3. Build settings:
   - Build command: *(để trống)*
   - Build output directory: **`site`**
4. Deploy. Site chạy hoàn toàn static — không tốn server, không cần cấu hình gì thêm.

### Chạy thử local

```bash
cd site && python3 -m http.server 8788
```

Mở http://localhost:8788.

## Cấu trúc

```
app/main.py     FastAPI app (bản local): API + serve UI
app/core.py     Logic Python: parse tên sheet, map cột, validate, ghi template
app/gsheet.py   Google OAuth + Sheets API (gspread) — chỉ bản local
static/         Giao diện bản local
template/       File template BigSeller (.xlsx) — bản local tự load file mới nhất
uploads/        File upload tạm (tạo tự động)
site/           Bản static cho Cloudflare Pages (index.html, core.js, app.js, template.xlsx)
```

## Tùy chỉnh

- Tên cột trong template: dict `TEMPLATE_FIELDS` ở đầu `app/core.py`.
- Cột nguồn trong Google Sheet: tìm theo tên header `BS - SKU` / `BS - Số lượng`
  (15 row đầu), không phụ thuộc vị trí cột.
- Quy tắc đặt tên đơn: `parse_sheet_name()` trong `app/core.py`.
