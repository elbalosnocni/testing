PAYROLLSYNC - NANG CAP TU KHOI PHUC MAT KHAU BANG SMS OTP
========================================================

CAC THAY DOI
------------
1. Sheet Excel DSCNV: doc so dien thoai tu cot O (15), tieu de
   "SO DIEN THOAI (Phone No.)".
2. VBA chuan hoa so dien thoai thanh chuoi 10 chu so:
   - Bo dau cham, khoang trang, gach ngang va ky tu khong phai so.
   - Neu Excel luu so dang number va mat so 0 dau (con 9 chu so), tu them 0.
   - Neu khong hop le thi gui chuoi rong, khong tu doan so.
3. API Google Apps Script them cot PhoneNumber (cot L) vao Employees.
   Cot nay duoc dinh dang Text. Neu lan dong bo sau so dien thoai rong/khong
   hop le thi giu nguyen so dien thoai cu, khong ghi de bang rong.
4. Giao dien dang nhap co nut Quen mat khau. Nhan vien nhap ma nhan vien va
   so dien thoai. Neu khop voi Employees va tai khoan Active, server moi gui OTP.
5. Xac minh OTP qua eOTP. OTP do eOTP sinh/kiem tra; PayrollSync khong luu ma OTP.
   OTP co hieu luc theo cau hinh du an eOTP (mac dinh 300 giay). eOTP tu gioi han
   5 lan nhap sai; PayrollSync gioi han yeu cau gui lai 60 giay/lan va toi da 5 yeu cau/6 gio cho moi ma nhan vien.
6. Khi reset thanh cong, cap nhat PasswordHash/PasswordSalt, bo MustChangePassword
   va thu hoi cac session nhan vien dang hoat dong. Ghi AuditLogs.

CAU HINH EOTP
-------------
- Tao tai khoan va du an website tai https://eotp.com.vn/
- Lay API Key va Secret trong tab Tich hop cua du an.
- Trong Apps Script: Project Settings > Script Properties, them:
      EASYOTP_API_KEY    = <API key cua du an>
      EASYOTP_API_SECRET = <Secret key cua du an>
- Khong dat hai khoa nay trong index.html, VBA, Google Sheets hoac file cong khai.
- Du an eOTP can duoc duyet; goi API su dung channel=sms.
- Goi y test: bat Test mode tren dashboard eOTP va dung dung so dien thoai test
  do eOTP hien thi. Kiem tra gui va verify truoc khi mo cho nhan vien.
- Goi SMS that co the phat sinh phi khi vuot quota/goi mien phi. Kiem tra quota,
  kenh gui va gia hien hanh tren dashboard truoc khi trien khai rong.

TRIEN KHAI
----------
1. Sao luu project Apps Script va Google Sheet hien tai.
2. Mo Apps Script dang duoc web app PayrollSync su dung; thay noi dung Code.gs
   bang Code.gs trong goi nang cap.
3. Them EASYOTP_API_KEY va EASYOTP_API_SECRET vao Script Properties.
4. Deploy > Manage deployments > Edit deployment > New version > Deploy.
   Dung URL /exec hien tai; khong can sua API_URL trong VBA/index neu giu deployment.
5. Thay module PayrollSync.bas trong file Excel/VBA bang ban moi. Neu import module
   moi thay vi paste de len module cu, hay xoa/doi ten module cu de tranh trung ten.
6. Chay dong bo PayrollSync. API se tu tao header PhoneNumber o cot L Employees.
7. Kiem tra mot so nhan vien trong Employees: PhoneNumber phai la chuoi 10 so,
   bat dau bang 0. Thu cac truong hop Excel co dau cham/khoang trang.
8. Kiem tra OTP tren tai khoan eOTP test truoc khi cho nhan vien su dung.

LUU Y
-----
- Neu nhan vien khong co so dien thoai hop le trong Employees thi khong the tu reset;
  Admin van co the dung chuc nang reset du phong.
- Man hinh tra loi thong diep chung de tranh tiet lo ma nhan vien/so dien thoai co ton tai.
- Ma OTP khong duoc gui tu trinh duyet; Apps Script goi eOTP server-to-server.
- Gioi han 60 giay hien tai la theo ma nhan vien; can them CAPTCHA/gioi han theo IP neu
  he thong bi public va co nguy co bi spam.
- Sau khi cap nhat code, kiem tra va thay doi cac khoa/secret cu da tung nam trong source
  code goc (SYNC_API_KEY, SETUP_KEY, tai khoan Admin mac dinh neu van dang dung).
- Khong gui file co khoa API/Secret eOTP cho nguoi khac.
