BO GOOGLE APPS SCRIPT KIEM TRA CHAM CONG - V1

CAI DAT
1. Mo Google Sheets can dung.
2. Extensions > Apps Script, dan toan bo Code.gs vao editor va Save.
3. Quay lai Google Sheets, tai lai trang.
4. Menu Cham cong tu dong > 1. Tao / kiem tra cac tab; cap quyen khi duoc hoi.
5. Nhap log vao RAW_LOG, lich ca thuc te vao CA_THUC_TE, kiem tra RULES.

RAW_LOG - cot toi thieu:
Ten May Cham Cong | Ma Nhan Vien | Thoi Gian Bam Gio | Ghi Chu
Ma Nhan Vien chinh la ma AC/MSVT. Khong xoa log cu; dan them log moi o cuoi. Timestamp nen la DateTime hoac dd/MM/yyyy HH:mm:ss.

CA_THUC_TE:
Ngay | MSVT | Ho Ten | Gio Vao Ca | Gio Ra Ca | Loai Ca | Xac Nhan | Ghi Chu
Moi nhan vien/moi ngay mot dong. Dan YES vao Xac Nhan chi khi lich thuc te da duoc xac nhan. Ca qua dem: Gio Ra Ca nen nhap DateTime ngay hom sau neu co the. Neu de trong/chua xac nhan, script khong ket luan tre/som.

RULES:
TRE_SOM_NGUONG_PHUT=10; TRE_SOM_TRU_CONG_PHUT=30; TRE_SOM_LAM_TRON_TU_PHUT=40 (tren 40 phut tinh 60 phut).
AN_QUA_BIEN_BAN_TU_PHUT=5 (qua tren 5 phut lap bien ban); AN_QUA_TRU_CONG_TU_PHUT=10; AN_QUA_TRU_CO_BAN_PHUT=30; AN_QUA_LAM_TRON_TU_PHUT=45 (tren 45 phut tinh 60 phut).
Cac cua so giai lao va thoi luong cho phep nam trong WINDOW1..WINDOW4. Mac dinh la 08:30-13:00, 15:30-19:00, 20:30-23:00, 00:30-04:30; thoi luong mac dinh 30, 30, 30, 45 phut. Dieu chinh phu hop ca va quy dinh cong ty.

CHAY
Menu > 2. Chay kiem tra. Co the tao trigger moi gio trong menu.
Ket qua vao KET_QUA va XAC_MINH. Neu tab cu 'An qua gio', 'Tre som', 'Xac nhan loi' ton tai, script se bo sung dong theo khoa, khong xoa dong cu. Cac tab cu '24', 'QQVT', 'Xac nhan lam viec', 'Ko phep' duoc giu nguyen, khong tu ghi vao khi chua biet dung cau truc cot.

CAN LUU Y
- Khong co lich ca thuc te da xac nhan thi khong the ket luan tre/som an toan.
- Nhan dien cap giai lao la heuristic: hai lan quet lien tiep nam trong cung cua so. Can nguoi phu trach xac nhan cap quet dung la ra nghi/vao lai.
- Hai dong timestamp trung khop tuyet doi chi duoc bo qua khi tinh toan, log goc van con trong RAW_LOG.
- V1 chua tu dong biet nhan vien bi benh, ve som dot xuat hay huy tang ca; can cap nhat lich/ghi chu va xac nhan truoc khi chot cong.
- Truoc khi su dung de chot cong, hay chay thu mot thang va doi chieu bao cao da duyet. Kiem tra ky cau truc cac tab bao cao cu truoc khi dua vao quy trinh chinh thuc.
