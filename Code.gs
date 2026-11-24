/** LỊCH LÀM VIỆC - API V3
 * Chạy setupSystem('MAT_KHAU_ADMIN_MOI') một lần trong Apps Script trước khi dùng.
 */
const SPREADSHEET_ID = '1GSPgNkJGz2igyKwP4lrCB6oNiuGb8PSjPw8zrgcHMOI';
// Sheet log máy chấm công riêng; mặc định đọc tab đầu tiên. Có thể điền tên tab cụ thể nếu cần.
const ATTENDANCE_LOG_SPREADSHEET_ID = '1LshZZc_Pqf2hTq6CCC-6y69RjOPZoylYbGroaAqUrNM';
const ATTENDANCE_LOG_SHEET_NAME = ''; // Để trống = dùng tab đầu tiên trong file log.
const LOG_CODE_COL = 3; // C: Mã chấm công
const LOG_TIME_COL = 4; // D: Thời gian bấm công
const SH = {
  schedule: ['Tháng','Mã NV','Mã chấm công','Họ tên','Phòng ban','Bộ phận','Chức vụ','Ngày','Ca trực','Cập nhật lúc','Người cập nhật'],
  submissions: ['ID','Tháng','Bộ phận','Người gửi','Gửi lúc','Trạng thái','Dữ liệu JSON','Duyệt lúc','Ghi chú'],
  shifts: ['Mã ca','Tên ca','Giờ bắt đầu','Giờ kết thúc','Màu sắc','Đang sử dụng'],
  depts: ['Bộ phận','Mật khẩu','Đang sử dụng','Cập nhật lúc'],
  audit: ['Thời gian','Vai trò','Bộ phận','Hành động','Chi tiết'],
  attendance: ['Thời gian xuất','Tháng','Mã NV','Mã chấm công','Họ tên','Phòng ban','Bộ phận','Chức vụ','Ngày','Ca trực','Loại ngày đã xác nhận','Tên ngày','Giờ vào thực tế','Giờ ra thực tế','Số lần bấm','Số giờ giữa vào/ra','Hệ số tham chiếu','Trạng thái tính lương','Trạng thái đối chiếu','Chi tiết log'],
  calendar: ['Ngày','Loại ngày','Tên ngày / lý do','Lịch điều chỉnh','Đã xác nhận','Người xác nhận','Cập nhật lúc','Ghi chú'],
  employees: ['Hoten','MaNV','Machacong','SoCancuoc','Phongban','Bophan','Chucvu','Totruong'],
  mapping: ['Bộ phận','Ký hiệu (chuẩn hóa)','Hiển thị gốc','Mã ca chung','Cập nhật lúc','Người cập nhật'],
  changes: ['Thời gian','Lô','Tháng','Bộ phận','Khóa NV','Mã NV','Mã chấm công','Họ tên','Ngày','Ca cũ','Ca mới','Ghi chú'],
  subs: ['ID băm','Khóa NV','Mã NV','Mã chấm công','Họ tên','Token FCM','Thiết bị','Tạo lúc','Lần cuối','Đang dùng'],
  blob: ['ID','Thứ tự','Dữ liệu','Tạo lúc']
};
function doGet(e) {
  try {
    const p=(e&&e.parameter)||{}, ss=ss_();
    if(p.action==='schedule_get') return out_(getSchedule_(ss,p.month));
    if(p.action==='calendar_get') return out_(getCalendar_(ss,p.month));
    if(p.action==='submission_status') return out_(getStatus_(ss,p.month));
    if(p.action==='get_shifts') return out_({status:'success',shifts:readShifts_(ss)});
    if(p.action==='health') return out_({status:'success',message:'LLV API V3 ready'});
    return out_({status:'error',message:'Action GET không được hỗ trợ.'});
  } catch(err){return out_({status:'error',message:String(err.message||err)});}
}
function doPost(e) {
  try {
    const b=JSON.parse((e&&e.postData&&e.postData.contents)||'{}'), ss=ss_();
    switch(b.action){
      case 'dept_login': {const dept=verifyDept_(ss,b.department,b.password);return out_(dept?{status:'success',department:dept}:{status:'error',message:'Sai bộ phận hoặc mật khẩu.'});}
      case 'admin_login': return out_(isAdmin_(b.adminPassword)?{status:'success'}:{status:'error',message:'Mật khẩu Admin không đúng.'});
      case 'hr_login': return out_(isHr_(b.hrPassword)?{status:'success'}:{status:'error',message:'Mật khẩu HR không đúng.'});
      case 'schedule_submit': return out_(submitSchedule_(ss,b));
      case 'department_approve': return out_(approveDepartment_(ss,b));
      case 'admin_get_config': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_({status:'success',departments:readDepartments_(ss),shifts:readShifts_(ss)});
      case 'admin_save_department': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_(saveDepartment_(ss,b));
      case 'admin_delete_department': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_(deleteDepartment_(ss,b));
      case 'admin_save_shifts': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_(saveShifts_(ss,b.shifts));
      case 'admin_save_calendar': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_(saveCalendar_(ss,b.month,b.days,b.adminName||'ADMIN')); 
      case 'attendance_export': if(!isAdmin_(b.adminPassword)&&!isHr_(b.hrPassword))return out_(denied_());return out_(exportAttendance_(ss,b.month));
      case 'mapping_get': return out_(mappingGet_(ss,b));
      case 'mapping_save': return out_(mappingSave_(ss,b));
      case 'dept_employees': return out_(deptEmployees_(ss,b));
      case 'sub_register': return out_(subRegister_(ss,b));
      case 'sub_token': return out_(subToken_(ss,b));
      case 'sub_unregister': return out_(subUnregister_(ss,b));
      case 'my_changes': return out_(myChanges_(ss,b));
      case 'admin_import_employees': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_(adminImportEmployees_(ss,b));
      case 'admin_push_test': if(!isAdmin_(b.adminPassword))return out_(denied_());return out_(adminPushTest_(ss,b));
      default:return out_({status:'error',message:'Action không được hỗ trợ.'});
    }
  }catch(err){return out_({status:'error',message:String(err.message||err)});}
}
function setupSystem(adminPassword, hrPassword){
  if(!adminPassword||String(adminPassword).length<10)throw new Error('Mật khẩu Admin phải có ít nhất 10 ký tự.');
  if(!hrPassword||String(hrPassword).length<10)throw new Error('Mật khẩu HR phải có ít nhất 10 ký tự.');
  const ss=ss_(); Object.keys(SH).forEach(k=>ensure_(ss,({schedule:'LICH_THANG',submissions:'LICH_GUI_DUYET',shifts:'CA_KIP',depts:'BO_PHAN',audit:'NHAT_KY',attendance:'XUAT_CHAM_CONG',calendar:'LICH_NGAY',employees:'NHAN VIEN',mapping:'MA_CA_BO_PHAN',changes:'LICH_THAY_DOI',subs:'THUE_BAO',blob:'DU_LIEU_TAM'})[k],SH[k]));
  const props=PropertiesService.getScriptProperties();props.setProperty('ADMIN_PASSWORD_HASH',hash_(String(adminPassword)));props.setProperty('HR_PASSWORD_HASH',hash_(String(hrPassword)));
  const shift=ss.getSheetByName('CA_KIP');if(shift.getLastRow()<2)shift.getRange(2,1,6,6).setValues([
    ['C1','Ca 1','06:00','14:00','#dcfce7',true],['C2','Ca 2','14:00','22:00','#fef3c7',true],['C3','Ca 3','22:00','06:00','#fee2e2',true],['HC','Hành chính','08:00','17:00','#dbeafe',true],['OFF','Nghỉ','00:00','00:00','#e2e8f0',true],['AL','Nghỉ phép','00:00','00:00','#f3e8ff',true]
  ]);
  Object.keys(SH).forEach(k=>{const n=({schedule:'LICH_THANG',submissions:'LICH_GUI_DUYET',shifts:'CA_KIP',depts:'BO_PHAN',audit:'NHAT_KY',attendance:'XUAT_CHAM_CONG',calendar:'LICH_NGAY',employees:'NHAN VIEN',mapping:'MA_CA_BO_PHAN',changes:'LICH_THAY_DOI',subs:'THUE_BAO',blob:'DU_LIEU_TAM'})[k];const s=ss.getSheetByName(n);s.setFrozenRows(1);s.getRange(1,1,1,s.getLastColumn()).setFontWeight('bold');});
  addDefaultShifts();
  return 'Khởi tạo hoàn tất. Hãy thêm bộ phận và mật khẩu trong trang admin.html.';
}
function getSchedule_(ss,month){if(!validMonth_(month))return {status:'error',message:'Tháng không hợp lệ.'};const s=ss.getSheetByName('LICH_THANG'),m={};if(s&&s.getLastRow()>=2){const v=s.getDataRange().getValues();for(let i=1;i<v.length;i++){const r=v[i];if(month_(r[0])!==month)continue;const key=String(r[2]||r[1]||r[3]);if(!m[key])m[key]={id:String(r[1]||''),maChamCong:String(r[2]||''),name:String(r[3]||''),dept:String(r[4]||''),section:String(r[5]||''),position:String(r[6]||''),shifts:{}};if(r[7]&&r[8]&&!isBlankOffShift_(r[8]))m[key].shifts[date_(r[7])]=String(r[8]);}}return {status:'success',month,employees:Object.values(m),shifts:readShifts_(ss).filter(x=>x.active),calendarDays:calendarMap_(ss,month)};}
function normalizeShiftKey_(value){return String(value==null?'':value).trim().toLocaleLowerCase('vi').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');}
function resolveShift_(value,shifts){
  const raw=normalizeShiftKey_(value);if(!raw)return null;
  for(const x of shifts){
    const code=normalizeShiftKey_(x.code),name=normalizeShiftKey_(x.name);
    if(raw===code||raw===name)return x;
    // Accept common schedule values "1", "2", "3" for configured names "Ca 1", "Ca 2", "Ca 3".
    const digits=raw.match(/^(?:ca\s*)?(\d+)$/),nameDigits=name.match(/^ca\s*(\d+)$/),codeDigits=code.match(/^c(\d+)$/);
    if(digits&&((nameDigits&&digits[1]===nameDigits[1])||(codeDigits&&digits[1]===codeDigits[1])))return x;
  }
  return null;
}
// ============================================================================
// V3 — Quy đổi ký hiệu theo bộ phận · so sánh thay đổi · nhật ký · thông báo đẩy
// ============================================================================
const SCRIPT_TZ_ = () => Session.getScriptTimeZone();
const MAX_CELL_ = 45000;                 // Google Sheets giới hạn 50.000 ký tự/ô
const NOTIFY_FROM_TODAY_ = true;         // chỉ báo thay đổi của ngày hôm nay trở đi

function withLock_(fn){const lock=LockService.getScriptLock();lock.waitLock(30000);try{return fn();}finally{lock.releaseLock();}}
function normCode_(v){let s=String(v==null?'':v).replace(/\s+/g,'').toUpperCase();if(/^\d+$/.test(s))s=s.replace(/^0+(?=\d)/,'');return s;}
function nameKey_(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9 ]/g,'').replace(/\s+/g,' ').trim();}
function pad2_(n){return (n<10?'0':'')+n;}
function empKey_(e){return String(e.maChamCong||e.id||e.name||'').trim();}
function today_(){return Utilities.formatDate(new Date(),SCRIPT_TZ_(),'yyyy-MM-dd');}
function fmtDM_(d){const p=String(d).split('-');return p.length===3?p[2]+'/'+p[1]:String(d);}

// ---------- Danh sách nhân viên gốc (tab NHAN VIEN) ----------
function readMaster_(ss){
  const s=ss.getSheetByName('NHAN VIEN');if(!s||s.getLastRow()<2)return [];
  return s.getDataRange().getDisplayValues().slice(1).filter(r=>String(r[0]||'').trim()).map(r=>({name:String(r[0]).trim(),id:String(r[1]||'').trim(),maChamCong:String(r[2]||'').trim(),cccd:String(r[3]||'').replace(/\D/g,''),dept:String(r[4]||'').trim(),section:String(r[5]||'').trim(),position:String(r[6]||'').trim(),leader:String(r[7]||'').trim()}));
}
function masterIndex_(master){
  const byCode={},byName={};
  master.forEach(m=>{[m.id,m.maChamCong].forEach(c=>{const k=normCode_(c);if(k){(byCode[k]=byCode[k]||[]).push(m);}});const nk=nameKey_(m.name);(byName[nk]=byName[nk]||[]).push(m);});
  return {byCode,byName};
}
/** Ghép dòng trong file với danh sách gốc: ưu tiên mã, sau đó tên (chỉ khi duy nhất). */
function matchMaster_(idx,e,dept){
  const codes=(Array.isArray(e.codes)?e.codes:[e.id,e.maChamCong]).concat([e.id,e.maChamCong]).map(normCode_).filter(Boolean);
  for(const c of codes){const hit=idx.byCode[c];if(hit&&hit.length){const same=hit.filter(m=>nameKey_(m.name)===nameKey_(e.name));return same[0]||hit[0];}}
  const byN=idx.byName[nameKey_(e.name)]||[];
  if(byN.length===1)return byN[0];
  if(byN.length>1){const d=byN.filter(m=>nameKey_(m.section)===nameKey_(dept));if(d.length===1)return d[0];}
  return null;
}

// ---------- Bảng quy đổi ký hiệu theo bộ phận ----------
function readMapping_(ss,dept){
  const s=ss.getSheetByName('MA_CA_BO_PHAN'),out={};if(!s||s.getLastRow()<2)return out;
  s.getDataRange().getValues().slice(1).forEach(r=>{if(String(r[0]).trim()===dept)out[normalizeShiftKey_(r[1])]={code:String(r[3]||''),raw:String(r[2]||r[1])};});
  return out;
}
function hourShiftCode_(start,end){
  const a=start.split(':'),b=end.split(':');
  const f=(p)=>Number(p[0])+'h'+(p[1]==='00'?'':p[1]);
  return f(a)+'-'+f(b); // ví dụ 7h-16h, 7h30-16h30 (tránh dạng 07-16 vì Excel tự đổi thành ngày)
}
const SHIFT_COLORS_=['#fde68a','#bbf7d0','#bfdbfe','#fbcfe8','#ddd6fe','#fed7aa','#a7f3d0','#fecaca','#c7d2fe','#fef08a','#99f6e4','#e9d5ff'];
/** Tạo ca chung theo khung giờ (mã ví dụ 7h-16h, 7h30-16h30). Không tạo trùng. */
function ensureHourShift_(ss,hours){
  const m=String(hours||'').match(/^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/);if(!m)throw new Error('Khung giờ không hợp lệ: '+hours);
  const start=pad2_(+m[1])+':'+m[2],end=pad2_(+m[3])+':'+m[4];
  if(+m[1]>23||+m[3]>23||+m[2]>59||+m[4]>59||start===end)throw new Error('Khung giờ không hợp lệ: '+hours);
  const code=hourShiftCode_(start,end),s=ensure_(ss,'CA_KIP',SH.shifts),v=s.getLastRow()>1?s.getDataRange().getValues():[];
  for(let i=1;i<v.length;i++)if(String(v[i][0]).trim().toUpperCase()===code){if(v[i][5]===false)s.getRange(i+1,6).setValue(true);return code;}
  const color=SHIFT_COLORS_[(s.getLastRow())%SHIFT_COLORS_.length];
  s.appendRow([code,'Ca '+start+'–'+end,start,end,color,true]);
  return code;
}
function mappingGet_(ss,b){
  const dept=verifyDept_(ss,b.department,b.password);if(!dept)return {status:'error',message:'Xác thực bộ phận thất bại.'};
  return {status:'success',department:dept,mappings:readMapping_(ss,dept),shifts:readShifts_(ss).filter(x=>x.active)};
}
function mappingSave_(ss,b){
  const dept=verifyDept_(ss,b.department,b.password);if(!dept)return {status:'error',message:'Xác thực bộ phận thất bại.'};
  if(!Array.isArray(b.items)||!b.items.length)return {status:'error',message:'Không có ký hiệu nào để lưu.'};
  return withLock_(()=>{
    const created=[],items=[];
    for(const it of b.items){
      const key=normalizeShiftKey_(it.key);if(!key)continue;
      let code=String(it.code||'').trim();
      if(code==='__NEW__'){code=ensureHourShift_(ss,it.newShift);created.push(code);}
      if(code!=='__IGNORE__'&&!readShifts_(ss).some(x=>x.code===code&&x.active))return {status:'error',message:'Mã ca "'+code+'" chưa có hoặc đã ngừng dùng (ký hiệu '+it.key+').'};
      items.push({key,raw:String(it.raw||it.key),code});
    }
    const s=ensure_(ss,'MA_CA_BO_PHAN',SH.mapping),v=s.getLastRow()>1?s.getDataRange().getValues():[SH.mapping],now=new Date(),pos={};
    for(let i=1;i<v.length;i++)if(String(v[i][0]).trim()===dept)pos[normalizeShiftKey_(v[i][1])]=i+1;
    items.forEach(x=>{const row=[dept,x.key,x.raw,x.code,now,dept];if(pos[x.key])s.getRange(pos[x.key],1,1,6).setValues([row]);else s.appendRow(row);});
    audit_(ss,'BỘ PHẬN',dept,'LƯU QUY ĐỔI KÝ HIỆU',items.length+' ký hiệu'+(created.length?' · ca mới: '+created.join(', '):''));
    return {status:'success',message:'Đã lưu '+items.length+' quy đổi'+(created.length?' và tạo ca chung mới: '+[...new Set(created)].join(', '):'')+'.',created:[...new Set(created)]};
  });
}

// ---------- Lưu dữ liệu tạm (tránh giới hạn 50.000 ký tự/ô) ----------
function saveBlob_(ss,id,obj){
  const s=ensure_(ss,'DU_LIEU_TAM',SH.blob),str=JSON.stringify(obj),rows=[];
  for(let i=0,n=0;i<str.length;i+=MAX_CELL_,n++)rows.push([id,n,str.slice(i,i+MAX_CELL_),new Date()]);
  s.getRange(s.getLastRow()+1,1,rows.length,4).setValues(rows);
  // dọn dữ liệu cũ hơn 14 ngày
  if(s.getLastRow()>1){const v=s.getRange(2,4,s.getLastRow()-1,1).getValues(),cut=Date.now()-14*864e5;let del=0;for(let i=0;i<v.length;i++){if(v[i][0] instanceof Date&&v[i][0].getTime()<cut)del++;else break;}if(del>0)s.deleteRows(2,del);}
}
function loadBlob_(ss,id){
  const s=ss.getSheetByName('DU_LIEU_TAM');if(!s||s.getLastRow()<2)return null;
  const v=s.getDataRange().getValues().slice(1).filter(r=>String(r[0])===id).sort((a,b)=>a[1]-b[1]);
  return v.length?JSON.parse(v.map(r=>String(r[2])).join('')):null;
}

// ---------- Chuẩn hóa dữ liệu gửi lên → mã ca chung ----------
/** Trả về {employees (ca đã đổi sang mã ca), unresolved:{khóa:{raw,count}}, errors, warnings, matched, unmatched} */
function prepareEmployees_(ss,dept,month,list){
  const shifts=readShifts_(ss),active=shifts.filter(x=>x.active),map=readMapping_(ss,dept);
  const master=readMaster_(ss),idx=masterIndex_(master),out=[],errors=[],warnings=[],unresolved={},seen={};
  const maxDay=new Date(Number(month.slice(0,4)),Number(month.slice(5,7)),0).getDate();
  let matched=0,unmatched=0,badFont=0;
  list.forEach((e,i)=>{
    const label='Dòng '+(e.row||i+1)+(e.name?' ('+e.name+')':'');
    const name=String(e.name||'').trim();if(!name){errors.push('Dòng '+(e.row||i+1)+' thiếu họ tên');return;}
    if(/[AEIOUY][ÃÕÙÛÖØÏÑÒ]/.test(name)&&name===name.toUpperCase())badFont++;
    const m=matchMaster_(idx,e,dept);
    const rec={id:String(e.id||'').trim(),maChamCong:String(e.maChamCong||'').trim(),name:name,dept:String(e.dept||e.group||'').trim(),section:dept,position:String(e.position||'').trim(),shifts:{}};
    if(m){matched++;rec.id=m.id||rec.id;rec.maChamCong=m.maChamCong||rec.maChamCong;rec.name=m.name;rec.position=rec.position||m.position;rec.dept=rec.dept||m.dept;
      if(m.section&&nameKey_(m.section)!==nameKey_(dept))warnings.push(label+' thuộc bộ phận "'+m.section+'" trong danh sách nhân viên (đang đăng nhập "'+dept+'").');}
    else{unmatched++;}
    if(!rec.id&&!rec.maChamCong&&!m){warnings.push(label+' không có mã và không tìm thấy theo tên — lịch được lưu theo tên, không gắn được thông báo.');}
    const ident=normCode_(rec.maChamCong||rec.id)||nameKey_(rec.name);

    Object.keys(e.shifts||{}).forEach(d=>{
      const day=Number(d.slice(-2));
      if(!/^\d{4}-\d{2}-\d{2}$/.test(d)||d.slice(0,7)!==month||day<1||day>maxDay){errors.push(label+' có ngày không hợp lệ '+d);return;}
      const raw=String(e.shifts[d]||'').trim();if(!raw)return;
      const mm=map[normalizeShiftKey_(raw)];
      if(mm&&mm.code==='__IGNORE__')return;
      let sh=mm?active.find(x=>x.code===mm.code):null;
      if(!sh&&!mm)sh=resolveShift_(raw,active);
      if(!sh){const k=normalizeShiftKey_(raw);unresolved[k]=unresolved[k]||{raw:raw,count:0};unresolved[k].count++;return;}
      rec.shifts[d]=sh.code;
    });
    if(seen[ident]){
      const prev=seen[ident];if(!Object.keys(prev.rec.shifts).length&&Object.keys(rec.shifts).length){out.splice(out.indexOf(prev.rec),1,rec);seen[ident]={row:e.row||i+1,rec:rec};return;}
      const same=JSON.stringify(prev.rec.shifts)===JSON.stringify(rec.shifts)||!Object.keys(rec.shifts).length;
      if(same){warnings.push(label+' trùng dòng '+prev.row+' (cùng nội dung) — đã bỏ qua dòng lặp.');return;}
      errors.push(label+' trùng với dòng '+prev.row+' nhưng ca khác nhau — cần xóa hoặc sửa 1 dòng');return;
    }
    seen[ident]={row:e.row||i+1,rec:rec};
    out.push(rec);
  });
  if(badFont>=3)warnings.push('Có '+badFont+' tên dùng font VNI/TCVN cũ (hiện ký tự lạ). Nên đổi sang Unicode để tra cứu và thông báo đúng người.');
  if(master.length&&unmatched)warnings.push(unmatched+' nhân viên không có trong danh sách nhân viên (tab NHAN VIEN) — vẫn lưu lịch nhưng không gắn được thông báo.');
  if(!master.length)warnings.push('Tab NHAN VIEN đang trống nên chưa đối chiếu được mã nhân viên. Admin nên nhập danh sách nhân viên để người lao động đăng ký nhận thông báo.');
  return {employees:out,unresolved,errors,warnings,matched,unmatched,total:list.length};
}

// ---------- Lịch đang công bố & so sánh ----------
function codeOfValue_(v,shifts){const raw=String(v||'').trim();if(!raw)return '';const k=normalizeShiftKey_(raw);const s=shifts.find(x=>normalizeShiftKey_(x.code)===k||normalizeShiftKey_(x.name)===k);return s?s.code:raw;}
function readPublished_(ss,month,dept){
  const s=ss.getSheetByName('LICH_THANG'),out={};if(!s||s.getLastRow()<2)return out;const shifts=readShifts_(ss);
  s.getDataRange().getValues().slice(1).forEach(r=>{
    if(month_(r[0])!==month||String(r[5]||'').trim()!==dept)return;
    const key=String(r[2]||r[1]||r[3]).trim();
    if(!out[key])out[key]={id:String(r[1]||''),maChamCong:String(r[2]||''),name:String(r[3]||''),shifts:{}};
    if(r[7]&&r[8])out[key].shifts[date_(r[7])]=codeOfValue_(r[8],shifts);
  });
  return out;
}
function diffSchedules_(oldMap,employees,mode){
  const changes=[],newEmployees=[],removed=[],seen={},st={employees:employees.length,newEmployees:0,removedEmployees:0,added:0,removedCells:0,changed:0,same:0};
  employees.forEach(e=>{
    const key=empKey_(e);seen[key]=1;const o=oldMap[key];
    if(!o){newEmployees.push({key,id:e.id,maChamCong:e.maChamCong,name:e.name});st.newEmployees++;return;}
    const dates={};Object.keys(o.shifts).forEach(d=>dates[d]=1);Object.keys(e.shifts).forEach(d=>dates[d]=1);
    Object.keys(dates).sort().forEach(d=>{
      const a=o.shifts[d]||'',b=e.shifts[d]||'';
      if(a===b){if(a)st.same++;return;}
      changes.push({key,id:e.id,maChamCong:e.maChamCong,name:e.name,date:d,oldV:a,newV:b});
      if(!a)st.added++;else if(!b)st.removedCells++;else st.changed++;
    });
  });
  if(mode==='replace')Object.keys(oldMap).forEach(key=>{if(!seen[key]){removed.push({key,name:oldMap[key].name});st.removedEmployees++;}});
  return {changes,newEmployees,removed,stats:st,mode};
}
function diffSummary_(d,limit){
  return {stats:d.stats,mode:d.mode,
    sample:d.changes.slice(0,limit||60).map(c=>({name:c.name,code:c.maChamCong||c.id,date:c.date,oldV:c.oldV,newV:c.newV})),
    moreChanges:Math.max(0,d.changes.length-(limit||60)),
    newEmployees:d.newEmployees.slice(0,30).map(x=>x.name),removed:d.removed.slice(0,30).map(x=>x.name)};
}

// ---------- Gửi / duyệt / công bố ----------
function submitSchedule_(ss,b){
  const dept=verifyDept_(ss,b.department,b.password);if(!dept)return {status:'error',message:'Xác thực bộ phận thất bại.'};
  if(!validMonth_(b.month))return {status:'error',message:'Tháng không hợp lệ.'};
  if(!Array.isArray(b.employees)||!b.employees.length)return {status:'error',message:'Chưa có danh sách nhân viên.'};
  const mode=b.mode==='replace'?'replace':'merge';
  const prep=prepareEmployees_(ss,dept,b.month,b.employees);
  const unresolvedKeys=Object.keys(prep.unresolved);
  if(unresolvedKeys.length)return {status:'error',code:'UNRESOLVED',message:'Còn '+unresolvedKeys.length+' ký hiệu chưa quy đổi sang mã ca chung: '+unresolvedKeys.slice(0,10).map(k=>prep.unresolved[k].raw).join(', ')+'.',unresolved:unresolvedKeys.map(k=>({key:k,raw:prep.unresolved[k].raw,count:prep.unresolved[k].count}))};
  if(prep.errors.length)return {status:'error',message:'Lịch chưa hợp lệ: '+prep.errors.slice(0,8).join('; ')};
  if(!prep.employees.length)return {status:'error',message:'Không có nhân viên hợp lệ.'};
  const diff=diffSchedules_(readPublished_(ss,b.month,dept),prep.employees,mode);
  const s=ensure_(ss,'LICH_GUI_DUYET',SH.submissions),id=Utilities.getUuid(),now=new Date();
  withLock_(()=>{
    saveBlob_(ss,id,{employees:prep.employees,mode:mode});
    s.appendRow([id,b.month,dept,dept,now,'Chờ bộ phận duyệt','BLOB:'+id,'','Chế độ: '+(mode==='replace'?'thay toàn bộ':'cập nhật theo nhân viên trong file')+' · '+prep.employees.length+' NV']);
    audit_(ss,'BỘ PHẬN',dept,'GỬI LỊCH',b.month+' / '+prep.employees.length+' nhân viên / '+mode);
  });
  return {status:'success',message:'Đã kiểm tra '+prep.employees.length+' nhân viên. Xem các thay đổi bên dưới rồi bấm Công bố.',submissionId:id,diff:diffSummary_(diff),warnings:prep.warnings,matched:prep.matched,unmatched:prep.unmatched};
}
function approveDepartment_(ss,b){
  const dept=verifyDept_(ss,b.department,b.password);if(!dept)return {status:'error',message:'Xác thực bộ phận thất bại.'};
  if(!validMonth_(b.month))return {status:'error',message:'Tháng không hợp lệ.'};
  const s=ss.getSheetByName('LICH_GUI_DUYET');
  if(!s||s.getLastRow()<2)return {status:'error',message:'Chưa có lịch nào được gửi. Hãy đọc file Excel rồi bấm “Kiểm tra & so sánh” trước.'};
  const v=s.getDataRange().getValues(),wanted=String(b.month).slice(0,7),nd=x=>String(x||'').trim().replace(/\s+/g,' ').toLocaleLowerCase('vi');
  let idx=-1;
  for(let i=v.length-1;i>=1;i--){const st=String(v[i][5]||'').trim().toLocaleLowerCase('vi');if(month_(v[i][1])===wanted&&nd(v[i][2])===nd(dept)&&(st==='chờ bộ phận duyệt'||st==='chờ bộ phận xác nhận'||st==='chờ xác nhận'||st==='pending')){idx=i+1;break;}}
  if(idx<0)return {status:'error',message:'Không có bản lịch đang chờ duyệt cho tháng '+wanted+'. Hãy bấm “Kiểm tra & so sánh” trước.'};
  const rec=v[idx-1];let payload;
  try{
    const ref=String(rec[6]||'');
    if(ref.indexOf('BLOB:')===0){payload=loadBlob_(ss,ref.slice(5));if(!payload)return {status:'error',message:'Dữ liệu tạm đã hết hạn. Hãy gửi kiểm tra lại.'};}
    else{const prep=prepareEmployees_(ss,dept,wanted,JSON.parse(ref||'[]'));if(Object.keys(prep.unresolved).length||prep.errors.length)return {status:'error',message:'Bản chờ cũ không còn hợp lệ, hãy gửi kiểm tra lại.'};payload={employees:prep.employees,mode:'replace'};}
  }catch(err){return {status:'error',message:'Dữ liệu lịch chờ bị lỗi: '+String(err.message||err)};}
  let res;
  try{res=withLock_(()=>publishDept_(ss,wanted,dept,payload.employees,payload.mode,dept));}
  catch(err){return {status:'error',message:'Không thể công bố lịch: '+String(err.message||err)};}
  if(res.rows<1)return {status:'error',message:'Không có ca hợp lệ để ghi cho tháng '+wanted+'. Lịch chưa được duyệt.'};
  s.getRange(idx,6).setValue('Đã duyệt bởi bộ phận');s.getRange(idx,8).setValue(new Date());s.getRange(idx,9).setValue('Đã ghi '+res.rows+' dòng · '+res.diff.stats.changed+' đổi ca, '+res.diff.stats.added+' thêm, '+res.diff.stats.removedCells+' bỏ · báo '+res.notify.sent+' thiết bị');
  audit_(ss,'BỘ PHẬN',dept,'DUYỆT & CÔNG BỐ',wanted+' / '+res.rows+' dòng / '+res.diff.changes.length+' thay đổi');
  return {status:'success',message:'Đã công bố lịch '+dept+' tháng '+wanted+' ('+res.rows+' dòng).'+(res.notify.sent?' Đã gửi thông báo tới '+res.notify.sent+' thiết bị.':''),rows:res.rows,diff:diffSummary_(res.diff,20),notify:res.notify};
}
/** mode 'merge': chỉ thay lịch của nhân viên có trong file; 'replace': thay toàn bộ lịch bộ phận trong tháng. */
function publishDept_(ss,month,dept,employees,mode,actor){
  const shifts=readShifts_(ss),now=new Date(),s=ensure_(ss,'LICH_THANG',SH.schedule);
  const oldMap=readPublished_(ss,month,dept),diff=diffSchedules_(oldMap,employees,mode);
  const rows=[];employees.forEach(e=>Object.keys(e.shifts||{}).forEach(d=>{if(e.shifts[d])rows.push([month,String(e.id||''),String(e.maChamCong||''),String(e.name||''),String(e.dept||''),dept,String(e.position||''),d,e.shifts[d],now,actor||dept]);}));
  if(!rows.length)return {rows:0,diff,notify:{sent:0}};
  const newKeys={};employees.forEach(e=>newKeys[empKey_(e)]=1);
  const v=s.getLastRow()>1?s.getDataRange().getValues():[],keep=[SH.schedule];
  for(let i=1;i<v.length;i++){
    const r=v[i].slice(0,SH.schedule.length);while(r.length<SH.schedule.length)r.push('');
    const target=month_(r[0])===month&&String(r[5]||'').trim()===dept;
    if(!target){keep.push(r);continue;}
    const key=String(r[2]||r[1]||r[3]).trim();
    if(mode!=='replace'&&!newKeys[key])keep.push(r); // giữ nguyên nhân viên không có trong file
  }
  const all=keep.concat(rows);
  s.clearContents();s.getRange(1,1,all.length,SH.schedule.length).setValues(all);s.setFrozenRows(1);
  SpreadsheetApp.flush();
  const batch=Utilities.getUuid().slice(0,8);
  writeChangeLog_(ss,batch,month,dept,diff,now);
  let notify={sent:0,skipped:'chưa cấu hình thông báo'};
  try{notify=notifyChanges_(ss,month,dept,diff,batch);}catch(err){notify={sent:0,error:String(err.message||err)};}
  return {rows:rows.length,diff,notify,batch};
}
function writeChangeLog_(ss,batch,month,dept,diff,now){
  const rows=[];
  diff.changes.slice(0,4000).forEach(c=>rows.push([now,batch,month,dept,c.key,c.id,c.maChamCong,c.name,c.date,c.oldV,c.newV,'']));
  diff.newEmployees.forEach(e=>rows.push([now,batch,month,dept,e.key,e.id,e.maChamCong,e.name,'','','LỊCH MỚI','']));
  if(!rows.length)return;
  const s=ensure_(ss,'LICH_THAY_DOI',SH.changes);s.getRange(s.getLastRow()+1,1,rows.length,SH.changes.length).setValues(rows);
  s.getRange(s.getLastRow()-rows.length+1,9,rows.length,1).setNumberFormat('@');
  // giữ nhật ký ~60 ngày
  const n=s.getLastRow();if(n>6000){const t=s.getRange(2,1,n-1,1).getValues(),cut=Date.now()-60*864e5;let del=0;for(let i=0;i<t.length;i++){if(t[i][0] instanceof Date&&t[i][0].getTime()<cut)del++;else break;}if(del>0)s.deleteRows(2,del);}
}

// ---------- Thuê bao (người lao động theo dõi lịch của mình) ----------
function subHash_(subId){return hash_('sub:'+String(subId||''));}
function readSubs_(ss){
  const s=ss.getSheetByName('THUE_BAO');if(!s||s.getLastRow()<2)return [];
  return s.getDataRange().getValues().slice(1).map((r,i)=>({row:i+2,hash:String(r[0]),key:String(r[1]),id:String(r[2]),maChamCong:String(r[3]),name:String(r[4]),token:String(r[5]||''),device:String(r[6]||''),active:r[9]!==false}));
}
function subRegister_(ss,b){
  const code=normCode_(b.code),last4=String(b.last4||'').replace(/\D/g,'').slice(-4),cache=CacheService.getScriptCache(),tk='subtry:'+code,tries=Number(cache.get(tk)||0);
  if(!code||last4.length<4)return {status:'error',message:'Nhập mã nhân viên/mã chấm công và 4 số cuối CCCD.'};
  if(tries>=8)return {status:'error',message:'Nhập sai quá nhiều lần. Vui lòng thử lại sau 1 giờ hoặc liên hệ nhân sự.'};
  const master=readMaster_(ss),m=master.filter(x=>normCode_(x.id)===code||normCode_(x.maChamCong)===code)[0];
  if(!m||!m.cccd||m.cccd.length<4||m.cccd.slice(-4)!==last4){
    cache.put(tk,String(tries+1),3600);
    return {status:'error',message:(m&&(!m.cccd||m.cccd.length<4))?'Danh sách nhân viên chưa có số CCCD của bạn. Vui lòng liên hệ nhân sự.':'Mã nhân viên hoặc 4 số cuối CCCD không đúng.'};
  }
  cache.remove(tk);
  return withLock_(()=>{
    const s=ensure_(ss,'THUE_BAO',SH.subs),key=String(m.maChamCong||m.id||m.name).trim(),now=new Date();
    let subId=String(b.subId||''),rows=readSubs_(ss),mine=subId?rows.filter(r=>r.hash===subHash_(subId))[0]:null;
    if(mine&&mine.key===key){s.getRange(mine.row,1,1,10).setValues([[mine.hash,key,m.id,m.maChamCong,m.name,mine.token,mine.device,now,now,true]]);}
    else{subId=Utilities.getUuid()+Utilities.getUuid().slice(0,8);s.appendRow([subHash_(subId),key,m.id,m.maChamCong,m.name,'',String(b.device||'').slice(0,120),now,now,true]);}
    return {status:'success',subId:subId,employee:{name:m.name,key:key,id:m.id,maChamCong:m.maChamCong,dept:m.section||m.dept,position:m.position}};
  });
}
function findSub_(ss,subId){const h=subHash_(subId);return readSubs_(ss).filter(r=>r.hash===h&&r.active)[0]||null;}
function subToken_(ss,b){
  const sub=findSub_(ss,b.subId);if(!sub)return {status:'error',code:'NO_SUB',message:'Phiên theo dõi không còn hiệu lực. Hãy đăng ký lại.'};
  const token=String(b.token||'').trim();if(token.length<20)return {status:'error',message:'Token thông báo không hợp lệ.'};
  return withLock_(()=>{
    const s=ss.getSheetByName('THUE_BAO');
    readSubs_(ss).forEach(r=>{if(r.token===token&&r.row!==sub.row)s.getRange(r.row,6).setValue('');}); // một thiết bị chỉ gắn một người
    s.getRange(sub.row,6).setValue(token);s.getRange(sub.row,7).setValue(String(b.device||sub.device).slice(0,120));s.getRange(sub.row,9).setValue(new Date());
    return {status:'success',message:'Đã bật thông báo trên thiết bị này.'};
  });
}
function subUnregister_(ss,b){
  const sub=findSub_(ss,b.subId);if(!sub)return {status:'success'};
  const s=ss.getSheetByName('THUE_BAO');s.getRange(sub.row,6).setValue('');s.getRange(sub.row,10).setValue(false);return {status:'success',message:'Đã ngừng theo dõi.'};
}
function myChanges_(ss,b){
  const sub=findSub_(ss,b.subId);if(!sub)return {status:'error',code:'NO_SUB',message:'Phiên theo dõi không còn hiệu lực.'};
  ss.getSheetByName('THUE_BAO').getRange(sub.row,9).setValue(new Date());
  const s=ss.getSheetByName('LICH_THAY_DOI'),out=[];
  if(s&&s.getLastRow()>1){
    const cut=Date.now()-45*864e5,v=s.getDataRange().getValues();
    for(let i=v.length-1;i>=1&&out.length<150;i--){
      const r=v[i];if(String(r[4])!==sub.key)continue;if(!(r[0] instanceof Date)||r[0].getTime()<cut)continue;
      out.push({time:Utilities.formatDate(r[0],SCRIPT_TZ_(),"yyyy-MM-dd'T'HH:mm:ssXXX"),month:String(r[2]),date:String(r[8]||''),oldV:String(r[9]||''),newV:String(r[10]||''),batch:String(r[1])});
    }
  }
  return {status:'success',employee:{name:sub.name,key:sub.key,id:sub.id,maChamCong:sub.maChamCong},changes:out,pushEnabled:!!sub.token,pushAvailable:pushConfigured_()};
}

// ---------- Thông báo đẩy (Firebase Cloud Messaging HTTP v1) ----------
function pushConfigured_(){const p=PropertiesService.getScriptProperties();return !!(p.getProperty('FCM_PROJECT_ID')&&p.getProperty('FCM_SERVICE_ACCOUNT'));}
function fcmAccessToken_(){
  const cache=CacheService.getScriptCache(),c=cache.get('fcm_access');if(c)return c;
  const sa=JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT')),now=Math.floor(Date.now()/1000);
  const enc=o=>Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/,'');
  const unsigned=enc({alg:'RS256',typ:'JWT'})+'.'+enc({iss:sa.client_email,scope:'https://www.googleapis.com/auth/firebase.messaging',aud:'https://oauth2.googleapis.com/token',iat:now,exp:now+3600});
  const sig=Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(unsigned,sa.private_key)).replace(/=+$/,'');
  const res=UrlFetchApp.fetch('https://oauth2.googleapis.com/token',{method:'post',payload:{grant_type:'urn:ietf:params:oauth:grant-type:jwt-bearer',assertion:unsigned+'.'+sig},muteHttpExceptions:true});
  const j=JSON.parse(res.getContentText());if(!j.access_token)throw new Error('Không lấy được token FCM: '+res.getContentText().slice(0,200));
  cache.put('fcm_access',j.access_token,3000);return j.access_token;
}
/** items: [{sub, title, body, url, tag}] — gửi song song, tự gỡ token đã hết hạn. */
function sendPushBatch_(ss,items){
  if(!items.length||!pushConfigured_())return {sent:0,failed:0};
  const project=PropertiesService.getScriptProperties().getProperty('FCM_PROJECT_ID'),auth='Bearer '+fcmAccessToken_();
  const reqs=items.map(it=>({url:'https://fcm.googleapis.com/v1/projects/'+project+'/messages:send',method:'post',contentType:'application/json',headers:{Authorization:auth},muteHttpExceptions:true,
    payload:JSON.stringify({message:{token:it.sub.token,data:{title:it.title,body:it.body,url:it.url||'',tag:it.tag||'llv'},webpush:{headers:{Urgency:'high',TTL:'86400'}}}})}));
  let sent=0,failed=0;const dead=[];
  for(let i=0;i<reqs.length;i+=50){
    const chunk=reqs.slice(i,i+50),res=UrlFetchApp.fetchAll(chunk);
    res.forEach((r,j)=>{const code=r.getResponseCode();if(code===200)sent++;else{failed++;const t=r.getContentText();if(code===404||/UNREGISTERED|INVALID_ARGUMENT/.test(t))dead.push(items[i+j].sub);}});
  }
  if(dead.length){const s=ss.getSheetByName('THUE_BAO');dead.forEach(d=>s.getRange(d.row,6).setValue(''));}
  return {sent:sent,failed:failed};
}
function notifyChanges_(ss,month,dept,diff,batch){
  if(!pushConfigured_())return {sent:0,skipped:'chưa cấu hình thông báo (FCM)'};
  const subs=readSubs_(ss).filter(r=>r.active&&r.token);if(!subs.length)return {sent:0,skipped:'chưa có người đăng ký'};
  const byKey={};subs.forEach(r=>(byKey[r.key]=byKey[r.key]||[]).push(r));
  const todayStr=today_(),per={};
  diff.changes.forEach(c=>{if(!byKey[c.key])return;if(NOTIFY_FROM_TODAY_&&c.date<todayStr)return;(per[c.key]=per[c.key]||[]).push(c);});
  const items=[],base='index.html?me=1&month='+month,mo=month.slice(5,7)+'/'+month.slice(0,4);
  Object.keys(per).forEach(key=>{
    const ch=per[key].sort((a,b)=>a.date<b.date?-1:1),show=ch.slice(0,3).map(c=>fmtDM_(c.date)+': '+(c.oldV||'—')+' → '+(c.newV||'nghỉ/bỏ')).join(' · ');
    const body=ch.length+' ngày thay đổi — '+show+(ch.length>3?' …':'');
    byKey[key].forEach(sub=>items.push({sub,title:'Lịch làm việc của bạn có thay đổi',body,url:base,tag:'llv-'+month}));
  });
  diff.newEmployees.forEach(e=>{if(!byKey[e.key]||per[e.key])return;byKey[e.key].forEach(sub=>items.push({sub,title:'Đã có lịch làm việc tháng '+mo,body:'Bộ phận '+dept+' vừa công bố lịch. Bấm để xem.',url:base,tag:'llv-'+month}));});
  const r=sendPushBatch_(ss,items);
  return {sent:r.sent,failed:r.failed,people:Object.keys(per).length};
}
function adminPushTest_(ss,b){
  if(!pushConfigured_())return {status:'error',message:'Chưa cấu hình FCM_PROJECT_ID / FCM_SERVICE_ACCOUNT trong Script properties.'};
  const k=normCode_(b.key),subs=readSubs_(ss).filter(r=>r.active&&r.token&&(normCode_(r.key)===k||normCode_(r.id)===k||normCode_(r.maChamCong)===k));
  if(!subs.length)return {status:'error',message:'Chưa có thiết bị nào bật thông báo cho mã này.'};
  const r=sendPushBatch_(ss,subs.map(sub=>({sub,title:'Thử thông báo lịch làm việc',body:'Nếu bạn thấy thông báo này, hệ thống đã hoạt động.',url:'index.html?me=1',tag:'llv-test'})));
  return {status:r.sent?'success':'error',message:'Đã gửi '+r.sent+' thiết bị, lỗi '+r.failed+'.'};
}

// ---------- Danh sách nhân viên: nhập từ Excel/CSV & mẫu theo bộ phận ----------
function adminImportEmployees_(ss,b){
  if(!Array.isArray(b.rows)||!b.rows.length)return {status:'error',message:'File không có dữ liệu.'};
  const rows=b.rows.map(r=>[r.Hoten,r.MaNV,r.Machacong,r.SoCancuoc,r.Phongban,r.Bophan,r.Chucvu,r.Totruong].map(x=>String(x==null?'':x).trim())).filter(r=>r[0]&&(r[1]||r[2]));
  if(!rows.length)return {status:'error',message:'Không có dòng hợp lệ (cần Họ tên và Mã NV hoặc Mã chấm công).'};
  return withLock_(()=>{
    const s=ensure_(ss,'NHAN VIEN',SH.employees);s.clearContents();s.getRange(1,1,1,SH.employees.length).setValues([SH.employees]);
    s.getRange(2,1,rows.length,8).setNumberFormat('@').setValues(rows);s.setFrozenRows(1);
    audit_(ss,'ADMIN','','NHẬP DANH SÁCH NHÂN VIÊN',rows.length+' người');
    return {status:'success',message:'Đã nhập '+rows.length+' nhân viên (thay toàn bộ danh sách cũ).'};
  });
}
function deptEmployees_(ss,b){
  const dept=verifyDept_(ss,b.department,b.password);if(!dept)return {status:'error',message:'Xác thực bộ phận thất bại.'};
  const k=nameKey_(dept);let list=readMaster_(ss).filter(m=>nameKey_(m.section)===k).map(m=>({id:m.id,maChamCong:m.maChamCong,name:m.name,position:m.position,group:m.dept}));
  if(!list.length&&validMonth_(b.month)){const pub=readPublished_(ss,b.month,dept);list=Object.keys(pub).map(key=>({id:pub[key].id,maChamCong:pub[key].maChamCong,name:pub[key].name,position:'',group:''}));}
  return {status:'success',department:dept,employees:list,shifts:readShifts_(ss).filter(x=>x.active)};
}
function addDefaultShifts(){
  // Chạy 1 lần trong trình soạn thảo để bổ sung các mã ca chung còn thiếu.
  const ss=ss_(),s=ensure_(ss,'CA_KIP',SH.shifts),have={};
  if(s.getLastRow()>1)s.getDataRange().getValues().slice(1).forEach(r=>have[String(r[0]).trim().toUpperCase()]=1);
  const defs=[['C1','Ca 1','06:00','14:00','#dcfce7'],['C2','Ca 2','14:00','22:00','#fef3c7'],['C3','Ca 3','22:00','06:00','#fee2e2'],['HC','Hành chính','08:00','17:00','#dbeafe'],['OFF','Nghỉ','00:00','00:00','#e2e8f0'],['AL','Nghỉ phép','00:00','00:00','#f3e8ff'],['LE','Nghỉ lễ/Tết','00:00','00:00','#fecdd3'],['TS','Nghỉ thai sản','00:00','00:00','#fbcfe8'],['WFH','Làm việc online','00:00','00:00','#cffafe']];
  const add=defs.filter(d=>!have[d[0]]);add.forEach(d=>s.appendRow([d[0],d[1],d[2],d[3],d[4],true]));
  return 'Đã thêm '+add.length+' mã ca: '+add.map(d=>d[0]).join(', ');
}
function authorizeNow(){UrlFetchApp.fetch('https://www.google.com',{muteHttpExceptions:true});LockService.getScriptLock();CacheService.getScriptCache();return 'OK';}

function saveDepartment_(ss,b){const dept=String(b.department||'').trim(),pw=String(b.password||'');if(!dept||!pw)return {status:'error',message:'Cần nhập tên bộ phận và mật khẩu.'};if(pw.length<6)return {status:'error',message:'Mật khẩu bộ phận tối thiểu 6 ký tự.'};const s=ensure_(ss,'BO_PHAN',SH.depts),v=s.getDataRange().getValues();let row=-1;for(let i=1;i<v.length;i++)if(String(v[i][0]).trim()===dept){row=i+1;break;}const vals=[dept,hash_(pw),true,new Date()];if(row>0)s.getRange(row,1,1,4).setValues([vals]);else s.appendRow(vals);audit_(ss,'ADMIN','', 'CẬP NHẬT BỘ PHẬN',dept);return {status:'success',message:'Đã lưu bộ phận '+dept};}
function deleteDepartment_(ss,b){const dept=String(b.department||'').trim(),s=ss.getSheetByName('BO_PHAN');if(s&&s.getLastRow()>1){const v=s.getDataRange().getValues();for(let i=v.length-1;i>=1;i--)if(String(v[i][0]).trim()===dept)s.deleteRow(i+1);}audit_(ss,'ADMIN','','NGỪNG BỘ PHẬN',dept);return {status:'success',message:'Đã xóa/ngừng bộ phận '+dept};}
function saveShifts_(ss,shifts){if(!Array.isArray(shifts))return {status:'error',message:'Danh sách ca không hợp lệ.'};const seen={};for(const x of shifts){if(!String(x.code||'').trim()||!String(x.name||'').trim())return {status:'error',message:'Mỗi ca cần mã và tên.'};const k=String(x.name).trim().toLowerCase();if(seen[k])return {status:'error',message:'Tên ca bị trùng: '+x.name};seen[k]=1;const code=String(x.code||'').trim().toUpperCase();const name=String(x.name||'').trim().toLowerCase();const isOff=['OFF','AL'].includes(code)||['nghỉ','nghi','nghỉ phép','nghi phep'].includes(name);const start=String(x.start||'').trim();const end=String(x.end||'').trim();if(isOff){if((start&&!/^([01]\d|2[0-3]):[0-5]\d$/.test(start))||(end&&!/^([01]\d|2[0-3]):[0-5]\d$/.test(end)))return {status:'error',message:'Nếu nhập giờ cho ca nghỉ, giờ phải theo HH:MM.'};}else if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(start)||!/^([01]\d|2[0-3]):[0-5]\d$/.test(end))return {status:'error',message:'Ca làm việc phải có giờ bắt đầu và kết thúc theo HH:MM. Ca Nghỉ/Nghỉ phép có thể để trống giờ.'};}const s=ensure_(ss,'CA_KIP',SH.shifts);s.clearContents();s.getRange(1,1,1,SH.shifts.length).setValues([SH.shifts]);if(shifts.length)s.getRange(2,1,shifts.length,6).setValues(shifts.map(x=>[x.code,x.name,x.start,x.end,x.color||'#dbeafe',x.active!==false]));s.setFrozenRows(1);audit_(ss,'ADMIN','','CẬP NHẬT CA KÍP',shifts.length+' ca');return {status:'success',message:'Đã lưu '+shifts.length+' ca kíp.'};}
function exportAttendance_(ss,month){
  if(!validMonth_(month))return {status:'error',message:'Tháng không hợp lệ.'};
  const data=getSchedule_(ss,month);if(data.status!=='success')return data;
  const logs=readAttendanceLogs_(month), shifts=readShifts_(ss), shiftMap={}, calendar=data.calendarDays||{};
  shifts.forEach(x=>{shiftMap[String(x.name||'').trim().toLowerCase()]=x;shiftMap[String(x.code||'').trim().toLowerCase()]=x;});
  const now=new Date(),rows=[];
  data.employees.forEach(e=>Object.keys(e.shifts||{}).sort().forEach(d=>{
    const shiftName=String(e.shifts[d]||'').trim(), shift=shiftMap[shiftName.toLowerCase()]||null;
    const code=normalizeClockCode_(e.maChamCong||e.id);
    const dayLogs=(logs.byCode[code]||[]);
    const pair=matchPunchesToShift_(dayLogs,d,shift);
    const day=calendar[d]||Object.assign(dayTypeInfo_('NORMAL',''),{approved:true});
    const factor=day.type==='NORMAL'?1:(day.approved?day.factor:'');
    const payNote=day.approved?day.payNote:(day.type==='NORMAL'?day.payNote:'Chưa xác nhận loại ngày; không dùng hệ số này để tính lương');
    rows.push([now,month,e.id,e.maChamCong,e.name,e.dept,e.section,e.position,d,shiftName,day.type,day.name||'',
      pair.inTime||'',pair.outTime||'',pair.count,pair.hours===null?'':pair.hours,factor,payNote,pair.status,pair.detail]);
  }));
  const s=ensure_(ss,'XUAT_CHAM_CONG',SH.attendance);s.clearContents();s.getRange(1,1,1,SH.attendance.length).setValues([SH.attendance]);
  if(rows.length){s.getRange(2,1,rows.length,SH.attendance.length).setValues(rows);s.getRange(2,13,rows.length,2).setNumberFormat('yyyy-mm-dd hh:mm:ss');s.getRange(2,16,rows.length,1).setNumberFormat('0.00');}
  audit_(ss,'HR/ADMIN','','ĐỐI CHIẾU LOG CHẤM CÔNG',month+' / '+rows.length+' dòng');
  return {status:'success',message:'Đã đối chiếu lịch với log máy chấm công và tạo bảng XUAT_CHAM_CONG.',rows:rows.length,logRows:logs.rows,sheet:'XUAT_CHAM_CONG',unmatched:logs.unmatched};
}
function readAttendanceLogs_(month){
  const out={byCode:{},rows:0,unmatched:0};
  const logSS=SpreadsheetApp.openById(ATTENDANCE_LOG_SPREADSHEET_ID);
  const sh=ATTENDANCE_LOG_SHEET_NAME?logSS.getSheetByName(ATTENDANCE_LOG_SHEET_NAME):logSS.getSheets()[0];
  if(!sh)throw new Error('Không tìm thấy tab log máy chấm công. Kiểm tra ATTENDANCE_LOG_SHEET_NAME.');
  const last=sh.getLastRow();if(last<2)return out;
  // Display values giữ nguyên số 0 đầu của mã chấm công; raw values dùng để đọc thời gian.
  const display=sh.getRange(1,1,last,Math.max(LOG_CODE_COL,LOG_TIME_COL)).getDisplayValues();
  const raw=sh.getRange(1,1,last,Math.max(LOG_CODE_COL,LOG_TIME_COL)).getValues();
  const startDate=new Date(month+'-01T00:00:00');startDate.setDate(startDate.getDate()-1);const start=Utilities.formatDate(startDate,Session.getScriptTimeZone(),'yyyy-MM-dd');const nextDate=new Date(nextMonth_(month)+'T00:00:00');nextDate.setDate(nextDate.getDate()+1);const next=Utilities.formatDate(nextDate,Session.getScriptTimeZone(),'yyyy-MM-dd');
  for(let i=1;i<last;i++){
    const code=normalizeClockCode_(display[i][LOG_CODE_COL-1]);
    const dt=parsePunchDate_(raw[i][LOG_TIME_COL-1],display[i][LOG_TIME_COL-1]);
    if(!code||!dt){out.unmatched++;continue;}
    const dateKey=Utilities.formatDate(dt,Session.getScriptTimeZone(),'yyyy-MM-dd');
    if(dateKey<start||dateKey>=next)continue;
    if(!out.byCode[code])out.byCode[code]=[];
    out.byCode[code].push(dt);out.rows++;
  }
  Object.keys(out.byCode).forEach(k=>out.byCode[k].sort((a,b)=>a-b));
  return out;
}
function matchPunchesToShift_(allLogs,dateStr,shift){
  if(!shift||isOffShift_(shift))return {inTime:'',outTime:'',count:0,hours:null,status:'Không tính công',detail:'Ca nghỉ/không tìm thấy cấu hình giờ ca'};
  const startMin=timeMinutes_(shift.start),endMin=timeMinutes_(shift.end);
  if(startMin===null||endMin===null)return {inTime:'',outTime:'',count:0,hours:null,status:'Thiếu cấu hình',detail:'Chưa cấu hình giờ bắt đầu/kết thúc ca'};
  const base=new Date(dateStr+'T00:00:00');
  const start=new Date(base.getTime()+startMin*60000),end=new Date(base.getTime()+endMin*60000);
  if(endMin<=startMin)end.setDate(end.getDate()+1); // ca qua nửa đêm
  // Cửa sổ nhận diện log: 2 giờ trước giờ vào đến 2 giờ sau giờ ra.
  const winStart=new Date(start.getTime()-2*3600000),winEnd=new Date(end.getTime()+2*3600000);
  const punches=allLogs.filter(t=>t>=winStart&&t<=winEnd);
  if(!punches.length)return {inTime:'',outTime:'',count:0,hours:null,status:'Thiếu log vào/ra',detail:'Không có log trong cửa sổ ca'};
  // Chọn log vào gần đầu ca nhất ở phía trước/đầu ca; log ra gần cuối ca nhất ở phía sau/đến cuối ca.
  let ins=punches.filter(t=>t<=new Date(start.getTime()+2*3600000));
  let outs=punches.filter(t=>t>=new Date(end.getTime()-2*3600000));
  let tin=ins.length?ins.reduce((best,t)=>Math.abs(t-start)<Math.abs(best-start)?t:best):null;
  let tout=outs.length?outs.reduce((best,t)=>Math.abs(t-end)<Math.abs(best-end)?t:best):null;
  if(tin&&tout&&tout<=tin)tout=null;
  const fmt=t=>t?Utilities.formatDate(t,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'):'';
  const hours=tin&&tout?Math.round(((tout-tin)/3600000)*100)/100:null;
  let status='Đủ log';if(!tin||!tout)status='Cần kiểm tra';
  const detail=(!tin?'Thiếu log vào; ':'')+(!tout?'Thiếu log ra; ':'')+(tin&&tout?'Đã ghép log gần giờ ca dự kiến':'');
  return {inTime:fmt(tin),outTime:fmt(tout),count:punches.length,hours,status,detail};
}
function parsePunchDate_(raw,display){
  if(raw instanceof Date&&!isNaN(raw.getTime()))return raw;
  if(typeof raw==='number'&&raw>0)return new Date(Math.round((raw-25569)*86400000));
  const s=String(display||raw||'').trim();if(!s)return null;
  let m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if(m)return new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(m[4]||0),Number(m[5]||0),Number(m[6]||0));
  const d=new Date(s);return isNaN(d.getTime())?null:d;
}
function normalizeClockCode_(v){return String(v==null?'':v).trim().replace(/\s+/g,'');}
function timeMinutes_(v){const m=String(v||'').match(/^(\d{1,2}):(\d{2})/);return m?Number(m[1])*60+Number(m[2]):null;}
function isOffShift_(s){const n=String(s.name||'').trim().toLowerCase(),c=String(s.code||'').trim().toUpperCase();return c==='OFF'||c==='AL'||c==='LE'||c==='TS'||['nghỉ','nghi','nghỉ phép','nghi phep','nghỉ lễ/tết','nghỉ thai sản'].includes(n);}
function isBlankOffShift_(v){const n=normalizeShiftKey_(v);return n==='off'||n==='nghi'||n==='off day';}
function nextMonth_(m){const p=m.split('-').map(Number);return Utilities.formatDate(new Date(p[0],p[1],1),Session.getScriptTimeZone(),'yyyy-MM-dd');}
function getCalendar_(ss,month){if(!validMonth_(month))return {status:'error',message:'Tháng không hợp lệ.'};return {status:'success',month,days:calendarMap_(ss,month)};}
function calendarMap_(ss,month){const out={};const sh=ss.getSheetByName('LICH_NGAY');if(!sh||sh.getLastRow()<2)return out;const v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++){const r=v[i],d=date_(r[0]);if(!d||d.slice(0,7)!==month)continue;const type=String(r[1]||'NORMAL').trim().toUpperCase();out[d]=Object.assign(dayTypeInfo_(type,String(r[2]||'')),{date:d,type:type,name:String(r[2]||dayTypeInfo_(type,'').name),adjustment:String(r[3]||''),approved:r[4]===true||String(r[4]).toLowerCase()==='true',approvedBy:String(r[5]||''),note:String(r[7]||'')});}return out;}
function dayTypeInfo_(type,name){const t=String(type||'NORMAL').toUpperCase();const labels={NORMAL:'Ngày làm việc bình thường',WEEKLY_OFF:'Nghỉ hàng tuần',LEGAL_HOLIDAY:'Lễ/Tết theo luật',COMPANY_OFF:'Nghỉ đặc biệt công ty',COMPANY_WORKDAY:'Ngày làm bù/đổi lịch',ANNUAL_LEAVE:'Nghỉ phép năm'};let factor=1,payNote='Hệ số tham chiếu 100%; kiểm tra giờ làm thêm/ban đêm riêng';if(t==='WEEKLY_OFF'){factor=2;payNote='Tham chiếu tối thiểu 200% nếu là ngày nghỉ hàng tuần hợp lệ';}else if(t==='LEGAL_HOLIDAY'){factor=3;payNote='Tham chiếu tối thiểu 300% tiền làm thêm; tiền lương ngày lễ và điều kiện khác xét riêng';}else if(t==='COMPANY_OFF'){factor=null;payNote='Chưa tự áp hệ số: cần xác định có đồng thời là ngày nghỉ hàng tuần/ngày lễ theo luật không';}else if(t==='ANNUAL_LEAVE'){factor=null;payNote='Nghỉ phép: không tính như ca làm; xử lý tiền phép theo quy định riêng';}return {type:t,name:name||labels[t]||labels.NORMAL,factor,payNote,label:labels[t]||labels.NORMAL};}
function saveCalendar_(ss,month,days,adminName){if(!validMonth_(month)||!Array.isArray(days))return {status:'error',message:'Tháng hoặc danh sách ngày không hợp lệ.'};const max=new Date(Number(month.slice(0,4)),Number(month.slice(5,7)),0).getDate(),seen={},valid=['NORMAL','WEEKLY_OFF','LEGAL_HOLIDAY','COMPANY_OFF','COMPANY_WORKDAY','ANNUAL_LEAVE'];for(const x of days){const d=String(x.date||'');if(!/^\d{4}-\d{2}-\d{2}$/.test(d)||d.slice(0,7)!==month||Number(d.slice(-2))<1||Number(d.slice(-2))>max)return {status:'error',message:'Ngày không hợp lệ: '+d};if(seen[d])return {status:'error',message:'Ngày bị lặp: '+d};seen[d]=true;if(!valid.includes(String(x.type||'NORMAL').toUpperCase()))return {status:'error',message:'Loại ngày không hợp lệ: '+x.type};}const sh=ensure_(ss,'LICH_NGAY',SH.calendar),old=sh.getDataRange().getValues(),keep=[SH.calendar];for(let i=1;i<old.length;i++)if(date_(old[i][0]).slice(0,7)!==month)keep.push(old[i].slice(0,SH.calendar.length));const now=new Date();days.forEach(x=>keep.push([x.date,String(x.type||'NORMAL').toUpperCase(),String(x.name||''),String(x.adjustment||''),x.approved===true, x.approved===true?adminName:'',now,String(x.note||'')]));sh.clearContents();sh.getRange(1,1,keep.length,SH.calendar.length).setValues(keep);sh.setFrozenRows(1);audit_(ss,'ADMIN','', 'CẬP NHẬT LỊCH NGÀY',month+' / '+days.length+' ngày');return {status:'success',message:'Đã lưu '+days.length+' ngày cho '+month+'. Loại ngày được lưu riêng với ca trực.'};}
function getStatus_(ss,month){if(!validMonth_(month))return {status:'error',message:'Tháng không hợp lệ.'};const s=ss.getSheetByName('LICH_GUI_DUYET');if(!s||s.getLastRow()<2)return {status:'success',submissions:[]};return {status:'success',submissions:s.getDataRange().getValues().slice(1).filter(r=>String(r[1])===month).map(r=>({month:r[1],department:r[2],submittedAt:r[4],status:r[5]}))};}
function readDepartments_(ss){const s=ss.getSheetByName('BO_PHAN');if(!s||s.getLastRow()<2)return [];return s.getDataRange().getValues().slice(1).filter(r=>r[0]&&r[2]!==false).map(r=>({department:String(r[0]),active:r[2]!==false}));}
function readShifts_(ss){const s=ss.getSheetByName('CA_KIP');if(!s||s.getLastRow()<2)return [];return s.getDataRange().getValues().slice(1).filter(r=>r[0]&&r[1]).map(r=>({code:String(r[0]),name:String(r[1]),start:String(r[2]),end:String(r[3]),color:String(r[4]||'#dbeafe'),active:r[5]!==false}));}
function verifyDept_(ss,dept,pw){const s=ss.getSheetByName('BO_PHAN');if(!s||!dept||!pw)return '';const v=s.getDataRange().getValues(),h=hash_(String(pw));for(let i=1;i<v.length;i++)if(String(v[i][0]).trim()===String(dept).trim()&&v[i][2]!==false&&String(v[i][1])===h)return String(v[i][0]).trim();return '';}
function isAdmin_(pw){const h=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');return !!h&&hash_(String(pw||''))===h;}
function isHr_(pw){const h=PropertiesService.getScriptProperties().getProperty('HR_PASSWORD_HASH');return !!h&&hash_(String(pw||''))===h;}
function hash_(s){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8);return bytes.map(b=>(b+256).toString(16).slice(-2)).join('');}
function audit_(ss,role,dept,action,detail){ensure_(ss,'NHAT_KY',SH.audit).appendRow([new Date(),role,dept,action,detail]);}
function ss_(){return SpreadsheetApp.openById(SPREADSHEET_ID)}
function ensure_(ss,name,headers){let s=ss.getSheetByName(name);if(!s)s=ss.insertSheet(name);if(s.getLastRow()===0||s.getRange(1,1).getValue()===''){s.getRange(1,1,1,headers.length).setValues([headers]);}else{const current=s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getDisplayValues()[0];if(current.length<headers.length)s.getRange(1,1,1,headers.length).setValues([headers]);}s.setFrozenRows(1);return s;}
function validMonth_(m){return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(m||''))}
function month_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM');return String(v||'').slice(0,7)}
function date_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v||'').slice(0,10)}
function out_(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON)}
function denied_(){return {status:'error',message:'Không có quyền Admin.'}}
