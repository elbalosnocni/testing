/* llv-me.js — "Lịch của tôi": đăng ký theo dõi bằng mã NV + 4 số cuối CCCD, bật thông báo đẩy, xem thay đổi gần đây. */
(function(){
const LS=window.localStorage,CFG=self.LLV_CONFIG||{},$$=id=>document.getElementById(id);
const get=(k)=>{try{return LS.getItem(k)}catch(e){return null}},set=(k,v)=>{try{LS.setItem(k,v)}catch(e){}},del=(k)=>{try{LS.removeItem(k)}catch(e){}};
const E=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let me=null,onlyMine=false,pushAvailable=false;
try{me=JSON.parse(get('llv_me')||'null')}catch(e){}
const subId=()=>get('llv_sub')||'';
async function api(body){const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(body)});return r.json()}
const isIOS=/iphone|ipad|ipod/i.test(navigator.userAgent),standalone=window.matchMedia('(display-mode: standalone)').matches||navigator.standalone;
const pushSupported=()=>('serviceWorker' in navigator)&&('PushManager' in window)&&('Notification' in window)&&CFG.FIREBASE&&CFG.FIREBASE.projectId&&CFG.VAPID_KEY;

/* ---------- Giao diện ---------- */
function shiftText(v){if(!v)return '<span class="text-slate-400">không có ca</span>';const s=(typeof shifts!=='undefined'?shifts:[]).find(x=>String(x.code).toLowerCase()===String(v).toLowerCase());return '<b>'+E(v)+'</b>'+(s&&s.start&&!(s.start==='00:00'&&s.end==='00:00')?' <span class="text-slate-500">('+E(s.start)+'–'+E(s.end)+')</span>':'')}
function box(){return $$('mePanel')}
function renderPanel(changes){
  const p=box();if(!p)return;
  if(!me){
    p.innerHTML=`<div class="font-semibold">Theo dõi lịch của tôi</div><p class="text-sm text-slate-600">Nhập mã nhân viên (hoặc mã chấm công) và 4 số cuối CCCD để xem nhanh lịch của bạn và nhận thông báo khi lịch thay đổi.</p>
    <div class="flex flex-wrap gap-2 items-end"><label class="text-sm">Mã nhân viên<input id="meCode" class="block border rounded-lg p-2 mt-1" autocomplete="off"></label><label class="text-sm">4 số cuối CCCD<input id="meLast4" inputmode="numeric" maxlength="4" class="block border rounded-lg p-2 mt-1 w-28" autocomplete="off"></label><button id="meGo" class="bg-blue-700 text-white rounded-lg px-4 py-2">Theo dõi</button></div><div id="meMsg" class="text-sm text-red-700"></div>`;
    $$('meGo').onclick=register;return;
  }
  const seen=get('llv_seen')||'',unseen=(changes||[]).filter(c=>c.time>seen).length;
  let push='';
  if(pushAvailable&&pushSupported()){
    push=(Notification.permission==='granted'&&get('llv_push')==='1')?'<span class="text-emerald-700 text-sm">🔔 Đã bật thông báo trên thiết bị này</span>':'<button id="mePush" class="border rounded-lg px-3 py-1 text-sm">🔔 Bật thông báo trên thiết bị này</button>';
  }else if(isIOS&&!standalone)push='<span class="text-sm text-amber-700">Để nhận thông báo trên iPhone/iPad: bấm Chia sẻ → “Thêm vào Màn hình chính”, mở từ biểu tượng đó rồi bật thông báo (iOS 16.4 trở lên).</span>';
  else if(!pushAvailable)push='<span class="text-xs text-slate-500">Thông báo đẩy chưa được bật ở máy chủ — thay đổi vẫn hiện ở danh sách bên dưới mỗi khi mở trang.</span>';
  const list=(changes||[]).slice(0,30).map(c=>{
    const when=new Date(c.time),ws=isNaN(when)?'':when.toLocaleString('vi-VN',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'});
    const txt=c.newV==='LỊCH MỚI'?'Đã công bố lịch tháng '+E(c.month.slice(5)+'/'+c.month.slice(0,4)):`Ngày <b>${E(c.date.slice(8)+'/'+c.date.slice(5,7))}</b>: ${shiftText(c.oldV)} → ${shiftText(c.newV)}`;
    return `<li class="py-1 border-b text-sm ${c.time>seen?'bg-yellow-50':''}"><span class="text-xs text-slate-400">${E(ws)}</span> ${txt}</li>`}).join('');
  p.innerHTML=`<div class="flex flex-wrap items-center gap-2"><div class="font-semibold">Lịch của ${E(me.name)}</div>${unseen?`<span class="bg-red-600 text-white text-xs rounded-full px-2 py-0.5">${unseen} thay đổi mới</span>`:''}
  <label class="text-sm ml-auto"><input type="checkbox" id="meOnly" ${onlyMine?'checked':''}> Chỉ hiện lịch của tôi</label><button id="meOut" class="text-xs text-slate-500 underline">Ngừng theo dõi</button></div>
  <div class="mt-1">${push}</div><div id="meMsg" class="text-sm text-red-700"></div>
  ${list?`<div class="mt-2 text-sm font-semibold">Thay đổi gần đây</div><ul class="max-h-52 overflow-auto">${list}</ul><button id="meSeen" class="text-xs text-blue-700 underline mt-1">Đánh dấu đã xem</button>`:'<div class="text-sm text-slate-500 mt-2">Chưa có thay đổi lịch nào trong 45 ngày qua.</div>'}`;
  $$('meOnly').onchange=e=>{onlyMine=e.target.checked;decorate()};
  $$('meOut').onclick=unregister;
  if($$('mePush'))$$('mePush').onclick=enablePush;
  if($$('meSeen'))$$('meSeen').onclick=()=>{set('llv_seen',new Date().toISOString());refresh()};
}
function setMsg(t){const m=$$('meMsg');if(m)m.textContent=t||''}

/* ---------- Đăng ký / huỷ ---------- */
async function register(){
  try{
    const code=$$('meCode').value.trim(),last4=$$('meLast4').value.trim();setMsg('');
    const j=await api({action:'sub_register',code,last4,subId:subId(),device:navigator.userAgent.slice(0,100)});
    if(j.status!=='success')throw Error(j.message);
    set('llv_sub',j.subId);me=j.employee;set('llv_me',JSON.stringify(me));set('llv_seen',new Date().toISOString());
    await refresh();decorate();
    if(pushSupported()&&pushAvailable&&(!isIOS||standalone))setTimeout(()=>{if(confirm('Bật thông báo để nhận tin ngay khi lịch của bạn thay đổi?'))enablePush()},200);
  }catch(e){setMsg(e.message)}
}
async function unregister(){
  if(!confirm('Ngừng theo dõi và tắt thông báo trên thiết bị này?'))return;
  try{await api({action:'sub_unregister',subId:subId()})}catch(e){}
  ['llv_sub','llv_me','llv_push','llv_seen'].forEach(del);me=null;onlyMine=false;renderPanel();decorate();
}

/* ---------- Thông báo đẩy ---------- */
function loadScript(src){return new Promise((ok,no)=>{const s=document.createElement('script');s.src=src;s.onload=ok;s.onerror=()=>no(Error('Không tải được '+src));document.head.appendChild(s)})}
async function enablePush(){
  try{
    setMsg('');
    if(!pushSupported())throw Error('Trình duyệt này chưa hỗ trợ thông báo đẩy.');
    const perm=await Notification.requestPermission();
    if(perm!=='granted')throw Error('Bạn chưa cho phép thông báo. Hãy bật quyền thông báo cho trang này trong cài đặt trình duyệt.');
    const reg=await navigator.serviceWorker.register('sw.js');await navigator.serviceWorker.ready;
    if(!window.firebase){await loadScript('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');await loadScript('https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js')}
    if(!firebase.apps.length)firebase.initializeApp(CFG.FIREBASE);
    const token=await firebase.messaging().getToken({vapidKey:CFG.VAPID_KEY,serviceWorkerRegistration:reg});
    if(!token)throw Error('Không lấy được mã thiết bị.');
    const j=await api({action:'sub_token',subId:subId(),token,device:navigator.userAgent.slice(0,100)});
    if(j.status!=='success')throw Error(j.message);
    set('llv_push','1');await refresh();
  }catch(e){setMsg(e.message)}
}

/* ---------- Dữ liệu & làm nổi bật dòng của tôi ---------- */
async function refresh(){
  if(!subId()||!me){renderPanel();return}
  try{
    const j=await api({action:'my_changes',subId:subId()});
    if(j.status!=='success'){if(j.code==='NO_SUB'){['llv_sub','llv_me','llv_push'].forEach(del);me=null;renderPanel();return}throw Error(j.message)}
    pushAvailable=!!j.pushAvailable;me=Object.assign(me,j.employee);renderPanel(j.changes);
  }catch(e){renderPanel([]);setMsg('Không tải được thay đổi: '+e.message)}
}
function decorate(){
  const t=$$('table');if(!t)return;
  const mc=me&&String(me.maChamCong||'').trim(),mid=me&&String(me.id||'').trim();
  t.querySelectorAll('tbody tr').forEach(tr=>{
    const mine=!!me&&((mc&&tr.children[1]&&tr.children[1].textContent.trim()===mc)||(mid&&tr.querySelector('.text-xs')&&tr.querySelector('.text-xs').textContent.trim().indexOf(mid+' ')===0));
    tr.classList.toggle('mine',mine);tr.style.display=(onlyMine&&me&&!mine)?'none':'';
  });
  const mine=t.querySelector('tr.mine');if(mine&&onlyMine)mine.scrollIntoView({block:'nearest'});
}
// gắn vào hàm render/load của index.html
const _render=window.render||render;
window.render=function(){
  // chỉ hiện trong chú thích những ca đang được dùng
  const all=shifts,used=new Set();employees.forEach(e=>Object.values(e.shifts||{}).forEach(v=>used.add(String(v).toLowerCase())));
  shifts=all.filter(s=>used.has(String(s.code).toLowerCase())||used.has(String(s.name).toLowerCase()));
  try{_render()}finally{shifts=all}
  decorate();
};
document.getElementById('q').oninput=window.render;document.getElementById('dept').onchange=window.render;
const style=document.createElement('style');style.textContent='.mine td{background:#fef9c3!important}.mine td.left{border-left:4px solid #ca8a04}';document.head.appendChild(style);

if('serviceWorker' in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('sw.js').catch(()=>{});
const q=new URLSearchParams(location.search);
if(q.get('month')&&/^\d{4}-\d{2}$/.test(q.get('month'))){document.getElementById('month').value=q.get('month');if(typeof load==='function')load()}
if(q.get('me')==='1'&&me){onlyMine=true}
renderPanel();refresh().then(()=>{if(onlyMine)decorate()});
})();
