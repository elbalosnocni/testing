PAYROLL V2 - KIEM TRA VA NANG CAP

1. Khoa nhan vien
- CitizenID/CCCD la khoa nghiep vu duy nhat.
- Chi chap nhan dung 12 chu so.
- Khong tu dong them so 0 cho CCCD sai do dai.
- EmployeeCode chi la ma nhan vien cua thang do.

2. Bao luu lich su
- Payroll luu theo tung SalaryMonth.
- Khi dong bo lai 08-2026, chi thay the du lieu 08-2026.
- 07-2026, 06-2026... khong bi xoa.
- Lich su tra cuu theo CitizenID, khong theo EmployeeCode.

3. Doi ma nhan vien
- Neu thang sau EmployeeCode thay doi, CCCD van la cung nguoi.
- Tai khoan Employees duoc cap nhat theo CCCD.
- Mat khau hien tai khong bi reset khi doi EmployeeCode.
- Lich su thang cu van hien thi.

4. Loc dong Excel
- Bat dau doc tu row 7.
- Gap TOTAL/TOTAL1/TOTAL2/TOTAL3... tai cot B thi dung doc.
- Chi gui dong co CCCD dung 12 so.
- Dong khong co CCCD hop le bi bo qua va ghi log.

5. MAP cot VBA
- Mo PayrollSync.bas.
- Chi can sua:
  PayrollColumnMap_()
  DSCNVColumnMap_()
- Vi du doi GrossIncome tu AR sang AT:
  m("grossIncome") = 46
- Khong sua CreatePayrollRecord_ hay JSON.

6. Mapping hien tai
Salary:
employeeCode D
basicSalary F
workingDays H
...
grossIncome AR
netSalary AW
fullName CF

DSCNV:
fullName B
citizenID H
department AF
section AG
position AH

7. GAS
- Code_v2.gs thay Code.gs.
- API van giu action login/getPayroll/syncPayroll va cac chuc nang admin.
- Session moi co them CitizenID de khong phu thuoc ma nhan vien.

8. Frontend
- index_v2.html thay index.html.
- Login bat buoc CCCD 12 so.

9. VBS
- RunPayrollSync_v2.vbs tu dong tim PayrollSync.xlsm trong cung thu muc voi VBS.
