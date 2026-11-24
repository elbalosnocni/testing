/* llv-ui.js — giao diện nhập lịch của bộ phận (Excel bộ phận / Excel chung) + tiện ích Admin.
 * Dùng các biến toàn cục của admin.html: $, post, msg, dept, password, adminPassword, API. */
(function(){
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const S={buf:null,sheets:[],parsed:{},catalog:[],saved:{},emps:[],symbols:{},legend:{},ym:null,submitted:false,mode:'dept'};
const ok=(t)=>msg(t,false),bad=(t)=>msg(t,true);
const skey=k=>LLV.strip(k);
function ym(){const m=$('month').value;if(!/^\d{4}-\d{2}$/.test(m))throw Error('Chọn tháng.');return {y:+m.slice(0,4),m:+m.slice(5,7),s:m};}

async function loadCatalogAndMapping(){
  const j=await post({action:'mapping_get',department:dept,password});
  if(j.status!=='success')throw Error(j.message);
  S.catalog=j.shifts;S.saved=j.mappings||{};
}
function pad(n){return String(n).padStart(2,'0')}

/* ---- 1. Chọn file → liệt kê sheet ---- */
async function onFile(){
  const f=$('file').files[0];if(!f)return;
  try{
    const t=ym();S.submitted=false;$('diffBox').innerHTML='';$('publish').disabled=true;
    S.buf=await f.arrayBuffer();
    const wb=XLSX.read(S.buf,{type:'array',bookSheets:true});
    const hidden=(wb.Workbook&&wb.Workbook.Sheets)||[];
    S.sheets=wb.SheetNames.map((n,i)=>({name:n,hidden:!!(hidden[i]&&hidden[i].Hidden),score:LLV.sheetScore(n,t.y,t.m)}));
    const vis=S.sheets.filter(x=>!x.hidden);
    let pick=vis.filter(x=>x.score>0);
    if(pick.length){const top=Math.max(...pick.map(x=>x.score));pick=pick.filter(x=>x.score===top);}
    else pick=vis.slice(0,12); // không có sheet theo tháng → thử các sheet đang hiện
    S.sheets.forEach(x=>x.use=pick.includes(x));
    $('sheetBox').innerHTML='<div class="text-sm font-semibold mb-1">Sheet dùng để đọc lịch (tích chọn lại nếu cần):</div><div class="flex flex-wrap gap-x-4 gap-y-1">'+S.sheets.map((x,i)=>`<label class="text-sm ${x.hidden?'text-slate-400':''}"><input type="checkbox" data-i="${i}" ${x.use?'checked':''}> ${esc(x.name.trim())}${x.hidden?' (ẩn)':''}${x.score?' ✓tháng':''}</label>`).join('')+'</div><button id="reparse" class="mt-2 border rounded-lg px-3 py-1 text-sm">Phân tích lại</button>';
    $('sheetBox').querySelectorAll('input').forEach(c=>c.onchange=()=>{S.sheets[+c.dataset.i].use=c.checked});
    $('reparse').onclick=()=>parseSelected().catch(e=>bad(e.message));
    await parseSelected();
  }catch(e){bad('Không đọc được file: '+e.message)}
}

/* ---- 2. Phân tích sheet đã chọn ---- */
async function parseSelected(){
  const t=ym(),names=S.sheets.filter(x=>x.use).map(x=>x.name);
  if(!names.length)throw Error('Chưa chọn sheet nào.');
  ok('Đang đọc '+names.length+' sheet…');await new Promise(r=>setTimeout(r,30));
  const wb=XLSX.read(S.buf,{type:'array',sheets:names});
  S.parsed={};const notes=[];
  names.forEach(n=>{
    const ws=wb.Sheets[n];if(!ws)return;
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:true});
    const r=LLV.parseSheet(rows,t.y,t.m,S.parsed[n]&&S.parsed[n].override);
    if(r){r.rows=rows;S.parsed[n]=r;}else notes.push(n.trim());
  });
  if(!Object.keys(S.parsed).length)throw Error('Không nhận dạng được bảng lịch trong các sheet đã chọn (cần dòng tiêu đề có các ngày 1…31 và cột Họ tên).');
  renderColumns(notes);
  await loadCatalogAndMapping();
  collect();renderMapping();
}
function colOptions(r,sel,allowNone){
  const o=allowNone?['<option value="-1">(không có)</option>']:[];
  for(let c=0;c<r.grid.firstCol;c++){const lab=(r.labels[c]||'').split(' / ').pop();o.push(`<option value="${c}" ${c===sel?'selected':''}>${LLV_col(c)}: ${esc(lab.slice(0,22)||'(không tên)')}</option>`)}
  return o.join('');
}
function LLV_col(c){let s='',n=c+1;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26)}return s}
function renderColumns(notes){
  const rows=Object.entries(S.parsed).map(([n,r])=>`<tr class="border-b" data-sheet="${esc(n)}"><td class="p-1 font-medium">${esc(n.trim())}</td><td class="p-1">${r.employees.length} NV</td>
  <td class="p-1"><select class="cs border rounded p-1" data-k="name">${colOptions(r,r.cols.name,false)}</select></td>
  <td class="p-1"><select class="cs border rounded p-1" data-k="id">${colOptions(r,r.cols.id,true)}</select></td>
  <td class="p-1"><select class="cs border rounded p-1" data-k="clock">${colOptions(r,r.cols.clock,true)}</select></td>
  <td class="p-1"><select class="cs border rounded p-1" data-k="pos">${colOptions(r,r.cols.pos,true)}</select></td>
  <td class="p-1"><select class="cs border rounded p-1" data-k="group">${colOptions(r,r.cols.group,true)}</select></td></tr>`).join('');
  $('colBox').innerHTML=`<div class="text-sm font-semibold mb-1">Cột thông tin (tự nhận dạng — sửa nếu sai):</div><div class="overflow-auto"><table class="text-sm"><thead><tr class="text-left border-b"><th class="p-1">Sheet</th><th class="p-1">Đọc được</th><th class="p-1">Họ tên</th><th class="p-1">Mã NV</th><th class="p-1">Mã chấm công</th><th class="p-1">Chức vụ</th><th class="p-1">Tổ/Nhóm</th></tr></thead><tbody>${rows}</tbody></table></div>`+(notes.length?`<div class="text-xs text-amber-700 mt-1">Không nhận dạng được bảng lịch: ${esc(notes.join(', '))}</div>`:'');
  $('colBox').querySelectorAll('select').forEach(sel=>sel.onchange=()=>{
    const n=sel.closest('tr').dataset.sheet,r=S.parsed[n],ov=Object.assign({},r.cols);
    sel.closest('tr').querySelectorAll('select').forEach(s=>ov[s.dataset.k]=+s.value);
    const t=ym(),np=LLV.parseSheet(r.rows,t.y,t.m,ov);if(np){np.rows=r.rows;np.override=ov;S.parsed[n]=np;sel.closest('tr').children[1].textContent=np.employees.length+' NV';collect();renderMapping();}
  });
}

/* ---- 3. Gộp nhân viên + ký hiệu ---- */
function collect(){
  const t=ym(),emps=[];S.symbols={};S.legend={};
  Object.entries(S.parsed).forEach(([n,r])=>{
    r.employees.forEach(e=>emps.push({row:e.row,sheet:n,name:e.name,id:e.id,maChamCong:e.clock,dept:e.group,position:e.position,
      shifts:Object.fromEntries(Object.entries(e.cells).map(([d,k])=>[`${t.s}-${pad(d)}`,k]))}));
    Object.entries(r.symbols).forEach(([k,v])=>{(S.symbols[k]=S.symbols[k]||{raw:v.raw,count:0}).count+=v.count});
    Object.assign(S.legend,r.legend);
  });
  S.emps=emps;S.sug=LLV.suggestMapping(S.symbols,S.legend,S.catalog);
  const warn=[];
  if(LLV.looksLegacyFont(emps.map(e=>e.name)))warn.push('Tên nhân viên có vẻ dùng font VNI/TCVN cũ (ký tự lạ). Nên chuyển sang Unicode trước khi gửi.');
  const nodays=emps.filter(e=>!Object.keys(e.shifts).length).length;if(nodays)warn.push(nodays+' dòng không có ca nào.');
  $('preview').innerHTML=`Đã đọc <b>${emps.length}</b> nhân viên, <b>${Object.keys(S.symbols).length}</b> ký hiệu ca khác nhau cho tháng ${t.s}. Bộ phận đăng nhập: <b>${esc(dept)}</b>.`+warn.map(w=>`<div class="text-amber-700">⚠ ${esc(w)}</div>`).join('');
}

/* ---- 4. Bảng quy đổi ký hiệu → mã ca chung ---- */
function codeOptions(selected,sug){
  let o='<option value="">— chọn mã ca —</option>';
  if(sug&&sug.code==='__NEW__'){const h=sug.newShift.split('-'),c=hourCode(sug.newShift);o+=`<option value="__NEW__" ${selected==='__NEW__'?'selected':''}>➕ Tạo ca chung mới ${h[0]}–${h[1]} (mã ${c})</option>`}
  o+=S.catalog.map(s=>{const h=s.start&&!(s.start==='00:00'&&s.end==='00:00')?` ${s.start}–${s.end}`:'';return `<option value="${esc(s.code)}" ${selected===s.code?'selected':''}>${esc(s.code)} · ${esc(s.name)}${h}</option>`}).join('');
  return o+`<option value="__IGNORE__" ${selected==='__IGNORE__'?'selected':''}>⊘ Bỏ qua (để trống ô này)</option>`;
}
function hourCode(h){const m=h.match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);if(!m)return h;const f=(x,y)=>(+x)+'h'+(y==='00'?'':y);return f(m[1],m[2])+'-'+f(m[3],m[4])}
function renderMapping(){
  const keys=Object.keys(S.symbols).sort((a,b)=>S.symbols[b].count-S.symbols[a].count);
  const rows=keys.map(k=>{
    const sv=S.saved[skey(k)],sug=S.sug[k],lg=S.legend[k];
    let sel='',tag='',cls='';
    if(sv){sel=sv.code;tag='<span class="text-emerald-700">đã lưu</span>'}
    else if(sug){sel=sug.code;tag=`<span class="${sug.conf==='low'?'text-amber-700':'text-blue-700'}">gợi ý: ${esc(sug.why)}</span>`;cls=sug.conf==='low'?'bg-amber-50':'bg-blue-50'}
    else{tag='<span class="text-red-700">chưa rõ — hãy chọn</span>';cls='bg-red-50'}
    const hint=lg?`<div class="text-xs text-slate-500">chú thích trong file: ${esc((lg.hours?lg.hours.replace('-','–')+' · ':'')+lg.text.slice(0,40))}</div>`:'';
    return `<tr class="border-b ${cls}" data-key="${esc(k)}"><td class="p-2 font-mono font-bold">${esc(S.symbols[k].raw)}</td><td class="p-2 text-right">${S.symbols[k].count}</td><td class="p-2"><select class="ms border rounded p-1 max-w-xs">${codeOptions(sel,sug)}</select>${hint}</td><td class="p-2 text-xs">${tag}</td></tr>`;
  }).join('');
  $('mapBox').innerHTML=`<div class="text-sm font-semibold mb-1">Quy đổi ký hiệu của bộ phận → mã ca chung</div><p class="text-xs text-slate-600 mb-2">Lần đầu hãy kiểm tra từng dòng. Quy đổi được lưu cho bộ phận ${esc(dept)}; các tháng sau tự áp dụng, chỉ hỏi lại ký hiệu mới.</p><div class="overflow-auto" style="max-height:340px"><table class="w-full text-sm"><thead class="sticky top-0 bg-white"><tr class="text-left border-b"><th class="p-2">Ký hiệu trong file</th><th class="p-2 text-right">Số ô</th><th class="p-2">Mã ca chung</th><th class="p-2">Nguồn</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}
function readMapping(){
  const items=[],missing=[];
  $('mapBox').querySelectorAll('tbody tr').forEach(tr=>{
    const k=tr.dataset.key,code=tr.querySelector('select').value;
    if(!code){missing.push(S.symbols[k].raw);return}
    const sv=S.saved[skey(k)],sug=S.sug[k];
    if(sv&&sv.code===code)return; // không đổi
    items.push({key:k,raw:S.symbols[k].raw,code,newShift:code==='__NEW__'?sug.newShift:undefined});
  });
  return {items,missing};
}

/* ---- 5. Kiểm tra & so sánh → Công bố ---- */
async function check(){
  try{
    if(!S.emps.length)throw Error('Chưa đọc file Excel.');
    const t=ym(),m=readMapping();
    if(m.missing.length)throw Error('Còn ký hiệu chưa chọn mã ca: '+m.missing.slice(0,8).join(', '));
    if(m.items.length){const j=await post({action:'mapping_save',department:dept,password,items:m.items});if(j.status!=='success')throw Error(j.message);await loadCatalogAndMapping();S.sug=LLV.suggestMapping(S.symbols,S.legend,S.catalog);renderMapping();}
    const mode=$('updMode').value;
    const j=await post({action:'schedule_submit',department:dept,password,month:t.s,employees:S.emps,mode});
    if(j.status!=='success'){if(j.code==='UNRESOLVED'){await loadCatalogAndMapping();renderMapping();}throw Error(j.message)}
    S.submitted=true;S.lastDiff=j.diff;$('publish').disabled=false;renderDiff(j);ok(j.message);
  }catch(e){bad(e.message)}
}
function renderDiff(j){
  const d=j.diff,s=d.stats,nm=x=>x||'—';
  const rows=d.sample.map(c=>`<tr class="border-b"><td class="p-1">${esc(c.name)}</td><td class="p-1 text-slate-500">${esc(c.code)}</td><td class="p-1">${esc(c.date.slice(8)+'/'+c.date.slice(5,7))}</td><td class="p-1">${esc(nm(c.oldV))} → <b>${esc(c.newV||'bỏ ca')}</b></td></tr>`).join('');
  $('diffBox').innerHTML=`<div class="border rounded-lg p-3 bg-slate-50 space-y-2"><div class="font-semibold">So sánh với lịch đang công bố</div>
  <div class="text-sm">${s.newEmployees?`<b>${s.newEmployees}</b> NV có lịch mới · `:''}<b>${s.changed}</b> ô đổi ca · <b>${s.added}</b> ô thêm · <b>${s.removedCells}</b> ô bỏ · ${s.same} ô giữ nguyên${s.removedEmployees?` · <span class="text-red-700"><b>${s.removedEmployees}</b> NV sẽ bị gỡ khỏi lịch tháng (chế độ thay toàn bộ)</span>`:''}</div>
  ${(j.warnings||[]).map(w=>`<div class="text-xs text-amber-700">⚠ ${esc(w)}</div>`).join('')}
  ${rows?`<div class="overflow-auto" style="max-height:240px"><table class="text-sm w-full"><thead><tr class="text-left border-b"><th class="p-1">Nhân viên</th><th class="p-1">Mã</th><th class="p-1">Ngày</th><th class="p-1">Thay đổi</th></tr></thead><tbody>${rows}</tbody></table></div>${d.moreChanges?`<div class="text-xs text-slate-500">…và ${d.moreChanges} thay đổi khác.</div>`:''}`:'<div class="text-sm text-slate-500">Không có thay đổi ca nào so với lịch đang công bố.</div>'}
  <div class="text-xs text-slate-500">Khớp danh sách nhân viên: ${j.matched}/${j.matched+j.unmatched}. Chỉ người đã bật thông báo và có thay đổi từ hôm nay trở đi mới được báo.</div></div>`;
}
async function publish(){
  try{
    if(!S.submitted)throw Error('Hãy bấm “Kiểm tra & so sánh” trước.');
    const rm=S.lastDiff.stats.removedEmployees;
    if(rm&&!confirm(`Chế độ “thay toàn bộ” sẽ GỠ ${rm} nhân viên không có trong file khỏi lịch tháng này. Tiếp tục?`))return;
    if(!confirm('Công bố lịch này cho nhân viên xem và gửi thông báo cho người có thay đổi?'))return;
    $('publish').disabled=true;
    const j=await post({action:'department_approve',department:dept,password,month:ym().s});
    if(j.status!=='success'){$('publish').disabled=false;throw Error(j.message)}
    S.submitted=false;ok(j.message);
  }catch(e){bad(e.message)}
}

/* ---- Mẫu Excel chung ---- */
async function downloadTemplate(){
  try{
    const t=ym(),j=await post({action:'dept_employees',department:dept,password,month:t.s});
    if(j.status!=='success')throw Error(j.message);
    const nd=LLV.daysIn(t.y,t.m),dow=['CN','T2','T3','T4','T5','T6','T7'];
    const head=['Mã NV','Mã chấm công','Họ tên','Chức vụ','Nhóm',...Array.from({length:nd},(_,i)=>i+1)];
    const wk=['','','','','',...Array.from({length:nd},(_,i)=>dow[new Date(t.y,t.m-1,i+1).getDay()])];
    const data=[[`LỊCH LÀM VIỆC ${dept.toUpperCase()} — THÁNG ${pad(t.m)}/${t.y} (điền MÃ CA CHUNG ở sheet MA_CA, để trống nếu không làm)`],[],wk,head];
    (j.employees.length?j.employees:[{}]).forEach(e=>data.push([e.id||'',e.maChamCong||'',e.name||'',e.position||'',e.group||'']));
    const ws=XLSX.utils.aoa_to_sheet(data);ws['!cols']=[{wch:10},{wch:13},{wch:28},{wch:16},{wch:12},...Array(nd).fill({wch:5})];ws['!freeze']={xSplit:3,ySplit:4};
    const ref=XLSX.utils.aoa_to_sheet([['Mã ca','Tên ca','Bắt đầu','Kết thúc'],...j.shifts.map(s=>[s.code,s.name,s.start,s.end])]);ref['!cols']=[{wch:10},{wch:26},{wch:9},{wch:9}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'LICH');XLSX.utils.book_append_sheet(wb,ref,'MA_CA');
    XLSX.writeFile(wb,`Mau_lich_${dept.replace(/\s+/g,'_')}_${t.s}.xlsx`);ok('Đã tạo mẫu với '+j.employees.length+' nhân viên của bộ phận.');
  }catch(e){bad(e.message)}
}

/* ---- Admin: nhập danh sách nhân viên & thử thông báo ---- */
async function importEmployees(){
  try{
    const f=$('empFile').files[0];if(!f)throw Error('Chọn file danh sách nhân viên (xlsx/csv).');
    const wb=XLSX.read(await f.arrayBuffer(),{type:'array'}),ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''}),H=LLV.strip;
    const hi=rows.findIndex(r=>r.some(x=>/ho ?ten|full ?name/.test(H(x))));
    const head=rows[hi<0?0:hi].map(x=>H(x).replace(/[^a-z0-9]/g,''));
    const col=(...n)=>head.findIndex(h=>n.includes(h));
    const ix={Hoten:col('hoten','hovaten','fullname','name'),MaNV:col('manv','msnv','manhanvien','employeeid'),Machacong:col('machacong','machamcong','mavantay','mvt'),SoCancuoc:col('socancuoc','cccd','socmnd','cancuoc'),Phongban:col('phongban','dept','department'),Bophan:col('bophan','section'),Chucvu:col('chucvu','position'),Totruong:col('totruong','leader')};
    if(ix.Hoten<0)throw Error('Không thấy cột Hoten/Họ tên.');
    const out=rows.slice((hi<0?0:hi)+1).map(r=>Object.fromEntries(Object.entries(ix).map(([k,i])=>[k,i>=0?r[i]:''])));
    const j=await post({action:'admin_import_employees',adminPassword,rows:out});if(j.status!=='success')throw Error(j.message);ok(j.message);
  }catch(e){bad(e.message)}
}
async function pushTest(){try{const j=await post({action:'admin_push_test',adminPassword,key:$('pushKey').value});j.status==='success'?ok(j.message):bad(j.message)}catch(e){bad(e.message)}}

function setMode(m){S.mode=m;$('modeNote').textContent=m==='dept'?'Tải đúng file Excel bộ phận đang dùng (mọi bố cục). Hệ thống tự tìm bảng, cột và ký hiệu ca.':'Dùng mẫu Excel chung: điền mã ca chung (xem sheet MA_CA). Hệ thống đọc trực tiếp, không cần quy đổi.';$('tplBtn').style.display=m==='common'?'inline-block':'none';$('modeDept').className='px-3 py-1 rounded-lg border '+(m==='dept'?'bg-slate-800 text-white':'');$('modeCommon').className='px-3 py-1 rounded-lg border '+(m==='common'?'bg-slate-800 text-white':'')}
window.addEventListener('DOMContentLoaded',()=>{
  $('file').onchange=onFile;$('month').onchange=()=>{if($('file').files[0])onFile()};
  $('check').onclick=check;$('publish').onclick=publish;$('tplBtn').onclick=downloadTemplate;
  $('modeDept').onclick=()=>setMode('dept');$('modeCommon').onclick=()=>setMode('common');setMode('dept');
  if($('empImport'))$('empImport').onclick=importEmployees;if($('pushTestBtn'))$('pushTestBtn').onclick=pushTest;
});
})();
