# Lịch làm việc V3 — đẩy lịch bằng Excel, quy mã ca chung, thông báo khi đổi lịch

## Có gì mới
| Việc | Cách làm trong V3 |
|---|---|
| Bộ phận đẩy lịch hàng tháng / mỗi lần đổi | `admin.html` → đăng nhập bộ phận → chọn tháng → chọn file Excel → **Kiểm tra & so sánh** → **Công bố** |
| Cách 1: Excel của bộ phận | Tự nhận dạng bảng, cột Họ tên/Mã, các ngày 1–31 và chọn đúng sheet của tháng. Đã thử trên cả 12 bộ phận trong `10_2026.rar` |
| Cách 2: Excel chung | Nút **Tải mẫu Excel chung** (đã điền sẵn nhân viên của bộ phận). Điền mã ca chung. Có sẵn file `Mau_Excel_chung.xlsx` (có danh sách chọn mã ca) |
| Mỗi bộ phận một ký hiệu | Bảng **Quy đổi ký hiệu → mã ca chung** theo từng bộ phận. Lần đầu bộ phận kiểm tra, lưu lại; các tháng sau tự áp dụng, chỉ hỏi ký hiệu mới |
| Gợi ý quy đổi | Đọc cả phần chú thích dưới bảng (vd. `X1: làm từ 7:00 đến 16:00`) để gợi ý. Chưa có ca đó → nút tạo ca chung theo giờ, mã dạng `7h-16h`, `7h30-16h30` |
| Cập nhật khi lịch đổi | Trước khi công bố hiện bảng **so sánh** (ai, ngày nào, ca cũ → ca mới). Có 2 chế độ: *chỉ cập nhật nhân viên có trong file* (mặc định, an toàn) hoặc *thay toàn bộ* |
| Nhật ký thay đổi | Tab `LICH_THAY_DOI` |
| Công nhân nhận thông báo | `index.html` → **Theo dõi lịch của tôi** (mã NV + 4 số cuối CCCD) → bật thông báo. Cài vào màn hình chính như một ứng dụng (PWA) |

Đã sửa lỗi cũ: lịch chờ duyệt lưu trong 1 ô nên bộ phận lớn (như Sản xuất ~330 người) vượt giới hạn 50.000 ký tự của Google Sheets. Nay lưu chia nhiều đoạn ở tab `DU_LIEU_TAM`.

## Cài đặt
1. **Sao lưu** Google Sheet và Apps Script hiện tại.
2. Apps Script: thay toàn bộ `Code.gs` bằng bản mới (giữ nguyên `SPREADSHEET_ID`, `ATTENDANCE_LOG_*`).
3. Chạy lần lượt trong trình soạn thảo: `addDefaultShifts` (thêm C1, C2, C3, HC, OFF, AL, LE, TS, WFH nếu thiếu), rồi `authorizeNow` (cấp quyền gọi dịch vụ ngoài cho thông báo). Các tab mới tự tạo khi dùng.
4. **Deploy → Manage deployments → Edit → New version → Deploy** (giữ URL cũ).
5. GitHub Pages: tải lên **tất cả** file trong gói (`index.html`, `admin.html`, `llv-parser.js`, `llv-ui.js`, `llv-me.js`, `config.js`, `sw.js`, `manifest.webmanifest`, `icon-*.png`). Nếu URL Apps Script đổi, sửa `API` ở đầu `index.html` và `admin.html`.
6. Admin → **Danh sách nhân viên & thông báo** → nhập file xuất từ phần mềm chấm công/SQL (cột `Hoten, MaNV, Machacong, SoCancuoc, Phongban, Bophan, Chucvu, Totruong`). Danh sách này dùng để ghép mã và xác minh công nhân khi đăng ký thông báo.
7. Admin → thêm bộ phận + mật khẩu (như cũ). Có thể đặt tên đúng như cột "Bộ phận" trong danh sách nhân viên.

## Bật thông báo đẩy (Firebase Cloud Messaging — miễn phí)
1. console.firebase.google.com → tạo project → thêm **Web app** → chép `firebaseConfig` vào `config.js`.
2. Project settings → Cloud Messaging → *Web Push certificates* → tạo Key pair → chép vào `VAPID_KEY` trong `config.js`.
3. Project settings → Service accounts → *Generate new private key* (file JSON). Trong Apps Script → Project Settings → **Script properties** thêm:
   - `FCM_PROJECT_ID` = ID project Firebase
   - `FCM_SERVICE_ACCOUNT` = dán nguyên nội dung file JSON
4. Đẩy lại `config.js` lên GitHub. Công nhân mở web → Theo dõi lịch của tôi → Bật thông báo. Admin có nút **Gửi thử** để kiểm tra.
- iPhone/iPad: cần iOS 16.4+, bấm Chia sẻ → *Thêm vào Màn hình chính*, mở từ biểu tượng đó rồi mới bật được thông báo.
- Chưa cấu hình FCM thì mọi thứ khác vẫn chạy; công nhân vẫn thấy danh sách "Thay đổi gần đây" mỗi lần mở trang.
- Chỉ báo những thay đổi của **hôm nay trở đi** cho người đã bật thông báo.

## Cách bộ phận dùng hàng tháng
1. Chọn tháng, chọn file Excel (cách 1) hoặc tải mẫu chung rồi điền (cách 2).
2. Kiểm tra sheet và cột được nhận dạng (sửa bằng danh sách chọn nếu sai).
3. Ở bảng quy đổi: dòng xanh = gợi ý, vàng = đoán (cần kiểm tra), đỏ = chưa rõ (phải chọn). Chọn *Bỏ qua* nếu ô đó không phải ca.
4. **Kiểm tra & so sánh** → xem thay đổi → **Công bố**.

## Kết quả thử trên file mẫu `10_2026.rar` (tháng 10/2026)
Nhận dạng đúng sheet và bảng của cả 12 bộ phận (kể cả file nhiều phiên bản `-2…-5`, dùng để thử so sánh thay đổi). Đã chạy thử toàn bộ luồng gửi → quy đổi → so sánh → công bố → nhật ký → đăng ký → thông báo bằng mô phỏng Apps Script với dữ liệu thật (Sản xuất 331 người, 9.900 dòng lịch).

**Việc cần bộ phận/nhân sự xử lý trong chính các file Excel:**
- **Kho NVL** gõ tên bằng font VNI cũ (`NGUYEÃN THÒ…`), không có cột mã → không ghép được nhân viên, không gắn được thông báo. Cần Unicode và thêm mã nhân viên.
- **Bếp, EHS+WW** không có cột mã; chỉ ghép theo tên (cần tên đúng như trong danh sách nhân viên).
- **Gia vị, QC4, PRO4** có dòng lặp (cùng người ghi 2 lần, dòng tên `NEW`, `Ms Truyền`). Dòng lặp cùng nội dung được bỏ qua; khác nội dung sẽ báo lỗi.
- **Chưa suy ra được ý nghĩa**, bộ phận phải tự chọn: Tài xế `6S`, `5S`, `7S`, `N`, `9:00`; QC `DS`, `NS`, `HC'`, `1-`, `2-`; Gia vị `29`, `4` (có thể là số gõ nhầm); PRO4 `5`.
- Cùng ký hiệu khác nghĩa giữa các bộ phận (vd. `A/B` PRO4 = 6h–18h/18h–6h, còn Gia vị `A` = 7h–16h) nên bảng quy đổi lưu **riêng theo bộ phận**.

## Chưa kiểm chứng được (cần bạn thử trên môi trường thật)
- Thông báo đẩy qua Firebase và PWA trên điện thoại: mình chỉ kiểm thử phần máy chủ bằng mô phỏng, chưa có Firebase/thiết bị thật.
- Giao diện `admin.html`/`index.html` mới: đã kiểm tra cú pháp và các id phần tử, bộ đọc Excel chạy trên dữ liệu thật, nhưng chưa mở trên trình duyệt thật.
- Xác minh người đăng ký bằng 4 số cuối CCCD là mức bảo vệ vừa phải (giới hạn 8 lần sai/giờ cho mỗi mã). Nếu cần chặt hơn nên thêm mã PIN do nhân sự cấp.
- `LICH_THANG` vẫn ghi lại toàn bộ tab mỗi lần công bố; nếu sau vài năm dữ liệu rất lớn nên lưu trữ các tháng cũ sang tab khác.
