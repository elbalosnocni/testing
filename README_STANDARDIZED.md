# Lịch làm việc — chuẩn hóa loại ngày và đối chiếu lương

## Điểm thay đổi
- Tạo tab `LICH_NGAY` để lưu loại ngày độc lập với ca trực: `NORMAL`, `WEEKLY_OFF`, `LEGAL_HOLIDAY`, `COMPANY_OFF`, `COMPANY_WORKDAY`, `ANNUAL_LEAVE`.
- Giao diện tra cứu giữ ô ngày nghỉ trống nhưng tô nền riêng; nếu có ca trực, ca vẫn hiện trên nền loại ngày gốc.
- Admin có thể quản lý lịch ngày theo tháng trong trang `admin.html`; cột “Đã xác nhận” dùng cho lịch điều chỉnh đã được phê duyệt.
- Xuất đối chiếu chấm công thêm loại ngày, tên ngày, hệ số tham chiếu và ghi chú tính lương.
- Lưu lịch ngày không thay đổi dữ liệu ca trong `LICH_THANG`; nhập ca không ghi đè `LICH_NGAY`.

## Cài đặt
1. Sao lưu Google Sheet và Apps Script.
2. Thay `Code.gs`, `index.html`, `admin.html` bằng các file trong gói. Giữ đúng `SPREADSHEET_ID` và URL API đang dùng.
3. Chạy lại `setupSystem` không bắt buộc nếu đã khởi tạo trước; tab `LICH_NGAY` sẽ được tạo tự động khi API được dùng hoặc khi setup chạy.
4. Deploy Apps Script: Deploy → Manage deployments → Edit → New version → Deploy. Cập nhật `index.html` và `admin.html` trên GitHub Pages.
5. Đăng nhập Admin → Lịch ngày nghỉ/đổi lịch → chọn tháng → thêm các ngày lễ, nghỉ tuần, nghỉ đặc biệt hoặc ngày làm bù → đánh dấu xác nhận khi đã được duyệt → Lưu.
6. Mở trang tra cứu, tải lại tháng. Màu ngày nghỉ chỉ phục vụ hiển thị; loại ngày lấy từ `LICH_NGAY`.

## Quy tắc tính lương / giới hạn quan trọng
- Hệ số trong `XUAT_CHAM_CONG` là tham chiếu: ngày thường 1.0, ngày nghỉ hàng tuần 2.0, lễ/Tết theo luật 3.0. Đây không phải bảng lương hoàn chỉnh và không tự tính phần lương ngày lễ, giờ làm thêm, giờ ban đêm hoặc các điều kiện pháp luật khác.
- `COMPANY_OFF` không tự gán 200%/300%; cần xác nhận ngày đó có đồng thời là ngày nghỉ hàng tuần hoặc ngày lễ theo luật hay không. `COMPANY_WORKDAY` dùng cho ngày đã điều chỉnh hợp lệ và không mặc nhiên là ngày nghỉ.
- Các ngày trong `LICH_NGAY` áp dụng toàn công ty. Nếu lịch nghỉ hàng tuần khác nhau theo từng nhân viên/bộ phận, cần mở rộng thành lịch ngày theo phạm vi nhân viên/bộ phận trước khi dùng kết quả để chốt lương.
- Đối chiếu cuối cùng cần HR/payroll xác nhận; không dùng màu ô làm căn cứ pháp lý.

## Kiểm tra sau triển khai
- Tab `LICH_NGAY` có tiêu đề cột và dữ liệu theo ngày.
- Đặt ca làm vào ngày `LEGAL_HOLIDAY` rồi tải lịch: ca vẫn hiện, nền ngày lễ vẫn giữ.
- Xuất chấm công: các cột loại ngày/hệ số tham chiếu phải có giá trị phù hợp.
- Kiểm tra quyền Admin và URL API trước khi dùng chính thức.


## Cập nhật giao diện v1.1
- Lịch tra cứu hiển thị **Mã ca** (ví dụ C1/C2/C3) thay cho tên ca; rê chuột lên mã ca để xem tên và giờ.
- Bảng chú thích đầu lịch lấy mã ca, tên ca, giờ bắt đầu/kết thúc và màu từ tab `CA_KIP`.
- Nền ngày được tô theo loại ngày trong `LICH_NGAY`: nghỉ tuần, lễ/Tết, nghỉ công ty, ngày làm bù và nghỉ phép.
- Xuất Excel bị chặn nếu người dùng đổi tháng nhưng chưa bấm **Tải lịch**, tránh tên file/tháng xuất không khớp tháng đang chọn.
- Khi công bố lịch mới, backend lưu mã ca thay vì tên ca. Dữ liệu lịch cũ đang lưu tên ca vẫn được giao diện nhận diện nếu tên trùng với `CA_KIP`.
- `setupSystem` tạo tab `NHAN VIEN` với bộ tiêu đề chuẩn nếu tab chưa có.

## Lưu ý phạm vi
Các yêu cầu lớn như nhập lịch trực tiếp trên web theo bộ phận/chức vụ/kíp/tổ trưởng, đồng bộ danh sách nhân viên mới nhất từ `NHAN VIEN`, và quy tắc chuyển bộ phận/tổ “cập nhật lịch hay giữ nguyên” chưa được triển khai trong bản giao diện này. Cần bổ sung API và màn hình quản trị tương ứng trước khi dùng thay thế hoàn toàn quy trình nhập lịch hiện tại.
