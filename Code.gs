// WIND booking v2. All scheduling uses Ho Chi Minh City time.
const SHEET_NAME = 'Lịch hẹn';
const NOTIFY_EMAIL = 'daoduykhuyen2@gmail.com';
const SALON_NAME = 'WIND HAIR HOUSE';
const SALON_PHONE = '098.186.6736';
const SALON_ADDR = 'Tòa S8.03, Vinhomes Grand Park, Thủ Đức, TP.HCM';
const FORM_SECRET = 'wind_d375b0682cd04279';
const TZ = 'Asia/Ho_Chi_Minh';
const CAPACITY = 3;
const STAFF = ['Phan Cường', 'Trung Ngọc', 'Đào Duy Khuyến'];
const DURATIONS = {'Cắt tóc':60,'Uốn tóc':240,'Duỗi tóc':240,'Nhuộm tóc':90,'Tẩy + Nhuộm':360,'Nối tóc':240};
const HEADERS = ['Thời gian gửi','Dịch vụ','Ngày hẹn','Khung giờ','Họ tên','Số điện thoại','Email','Lời nhắn','Trạng thái','Stylist','Thời lượng (phút)','Kết thúc','Mã lịch hẹn','Nhóm khách'];

function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function esc_(s) { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function safe_(s) { s=String(s || '').slice(0,500); return /^[=+@\-]/.test(s) ? "'"+s : s; }
function getSheet_() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}
function dateKey_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,TZ,'yyyy-MM-dd');
  var s=String(v || '').trim(), m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2) : s;
}
function timeKey_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,TZ,'HH:mm');
  var m=String(v || '').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  return m ? ('0'+m[1]).slice(-2)+':'+m[2] : '';
}
function start_(date,time) { return new Date(date+'T'+time+':00+07:00').getTime(); }
function duration_(s) {
  if(DURATIONS[s]) return DURATIONS[s];
  // Legacy rows keep their original values; reserve conservatively.
  if(/lông vũ|nối/i.test(s)) return 240;
  if(/tẩy|bóc|nâng sáng|highlight|babylight|balayage/i.test(s)) return 360;
  if(/uốn|duỗi/i.test(s)) return /nhuộm/i.test(s) ? 330 : 240;
  if(/nhuộm|phủ màu/i.test(s)) return 90;
  if(/cắt/i.test(s)) return 60;
  return 240;
}
function staffName_(s) {
  return /khuyến/i.test(s) ? STAFF[2] : /cường/i.test(s) ? STAFF[0] : /trung/i.test(s) ? STAFF[1] : '';
}
function rows_(sh) {
  if(sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,HEADERS.length).getValues().filter(function(r){return r[1] && !/^(đã hủy|đã huỷ|hủy|huỷ|cancelled|canceled|không đến|no.show)$/i.test(String(r[8]).trim());}).map(function(r){
    var begin=start_(dateKey_(r[2]),timeKey_(r[3]));
    if(!Number.isFinite(begin)) throw new Error('legacy_date_invalid');
    return {start:begin,end:begin+(Number(r[10])||duration_(String(r[1])))*60000,staff:staffName_(String(r[9])),id:String(r[12]||''),row:r};
  });
}
function peak_(rows,staff,a,b) {
  var events=[];
  rows.forEach(function(r){
    // Old unassigned bookings reserve capacity across every stylist until assigned.
    if((!r.staff || r.staff===staff) && r.start<b && r.end>a) {
      events.push([Math.max(a,r.start),1]); events.push([Math.min(b,r.end),-1]);
    }
  });
  events.sort(function(x,y){return x[0]-y[0] || x[1]-y[1];});
  var n=0,max=0;events.forEach(function(e){n+=e[1];max=Math.max(max,n);});return max;
}
function assign_(rows,a,b,requested,foreign) {
  var loads=STAFF.map(function(s){return peak_(rows,s,a,b);});
  var available=[0,1].filter(function(i){return loads[i]<CAPACITY;});
  var fallback=available.length===0;
  if(requested && requested!=='auto') {
    var i=STAFF.indexOf(requested);
    return i>=0 && loads[i]<CAPACITY && (i!==2 || foreign || fallback) ? STAFF[i] : '';
  }
  if(foreign && loads[2]<CAPACITY) return STAFF[2];
  available.sort(function(i,j){return loads[i]-loads[j] || i-j;});
  if(available.length) return STAFF[available[0]];
  return fallback && loads[2]<CAPACITY ? STAFF[2] : '';
}
function validSlot_(date,time,now) {
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||!/^\d{2}:[03]0$/.test(time)||time<'09:00'||time>'20:00')return false;
  var t=start_(date,time);
  return Number.isFinite(t) && Utilities.formatDate(new Date(t),TZ,'yyyy-MM-dd')===date && t>now && t<now+180*86400000;
}
function doGet(e) {
  try {
    var p=(e&&e.parameter)||{};
    if(p.action!=='availability')return json_({ok:true,version:2,msg:'WIND booking endpoint đang chạy.'});
    if(!DURATIONS[p.service] || !/^\d{4}-\d{2}-\d{2}$/.test(p.date||''))return json_({ok:false,error:'invalid_service_or_date'});
    var rows=rows_(getSheet_()),slots=[],now=Date.now();
    for(var m=540;m<=1200;m+=30) {
      var time=('0'+Math.floor(m/60)).slice(-2)+':'+('0'+m%60).slice(-2),a=start_(p.date,time);
      if(validSlot_(p.date,time,now)&&assign_(rows,a,a+DURATIONS[p.service]*60000,p.stylist||'auto',p.foreign==='true')) slots.push(time);
    }
    return json_({ok:true,version:2,slots:slots,duration:DURATIONS[p.service]});
  }catch(err){return json_({ok:false,error:String(err.message||err)});}
}
function doPost(e) {
  var lock=LockService.getScriptLock(),saved=false,d,id,assigned;
  try {
    d=JSON.parse(e.postData.contents);
    if(d.website)return json_({ok:false,error:'invalid_request'});
    if(d.token!==FORM_SECRET)return json_({ok:false,error:'unauthorized'});
    if(!d.name||String(d.name).trim().length<2||!d.phone||!DURATIONS[d.service]||!validSlot_(String(d.date),String(d.time),Date.now()))return json_({ok:false,error:'invalid_fields'});
    var phone=String(d.phone).replace(/[^0-9]/g,'');
    if(!/^\d{8,15}$/.test(phone))return json_({ok:false,error:'invalid_phone'});
    if(!/^[a-zA-Z0-9-]{16,80}$/.test(d.requestId||''))return json_({ok:false,error:'invalid_request_id'});
    if(d.email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email))return json_({ok:false,error:'invalid_email'});
    if(!lock.tryLock(25000))return json_({ok:false,error:'busy'});
    var sh=getSheet_(),rows=rows_(sh);
    var duplicate=rows.filter(function(r){return r.id===d.requestId;})[0];
    if(duplicate)return json_({ok:true,id:duplicate.id,stylist:duplicate.staff,date:dateKey_(duplicate.row[2]),time:timeKey_(duplicate.row[3]),service:String(duplicate.row[1])});
    var a=start_(d.date,d.time),b=a+DURATIONS[d.service]*60000;
    assigned=assign_(rows,a,b,d.stylist||'auto',d.foreign===true);
    if(!assigned)return json_({ok:false,error:'slot_full'});
    // A network retry with a new browser request ID must not create a duplicate.
    var same=rows.filter(function(r){return r.start===a&&String(r.row[5]).replace(/\D/g,'')===phone&&String(r.row[1])===d.service;})[0];
    if(same)return json_({ok:true,id:same.id||'existing',stylist:same.staff,date:d.date,time:d.time,service:d.service});
    id=d.requestId;
    sh.appendRow([Utilities.formatDate(new Date(),TZ,'dd/MM/yyyy HH:mm'),safe_(d.service),d.date,d.time,safe_(String(d.name).trim()),"'"+phone,safe_(d.email),safe_(d.note),'Đã xác nhận',assigned,DURATIONS[d.service],Utilities.formatDate(new Date(b),TZ,'yyyy-MM-dd HH:mm'),id,d.foreign===true?'Khách nước ngoài':'Khách Việt Nam']);
    SpreadsheetApp.flush();saved=true;
  }catch(err){return json_({ok:false,error:saved?'notification_error':'server_error'});}
  finally {if(lock.hasLock())lock.releaseLock();}
  // Email delivery cannot undo an appointment that has already been saved.
  try{notify_(d,id,assigned);}catch(err){console.error('Booking saved; email notification failed');}
  return json_({ok:true,id:id,stylist:assigned,date:d.date,time:d.time,service:d.service});
}
function notify_(d,id,stylist) {
  var detail='Mã lịch: '+id+'\nKhách: '+d.name+'\nSĐT: '+d.phone+'\nDịch vụ: '+d.service+'\nLịch: '+d.date+' '+d.time+'\nStylist: '+stylist+'\nLời nhắn: '+(d.note||'—');
  MailApp.sendEmail(NOTIFY_EMAIL,'Lịch hẹn đã xác nhận — '+d.name,detail+'\n\n'+SpreadsheetApp.getActiveSpreadsheet().getUrl());
  if(!d.email)return;
  var html='<div style="font-family:Arial;max-width:560px;margin:auto"><h2>WIND HAIR HOUSE</h2><p>Chào '+esc_(d.name)+',</p><p>WIND đã xác nhận và xếp lịch cho bạn. Vui lòng đến đúng giờ hẹn.</p><pre style="white-space:pre-wrap">'+esc_(detail)+'</pre><p>Cần đổi lịch, vui lòng <a href="tel:0981866736">gọi '+SALON_PHONE+'</a> hoặc liên hệ <a href="https://zalo.me/4548028938819960082">Zalo OA Wind Official</a>.</p><hr><p>Your appointment is confirmed. Please arrive on time. To reschedule, call the salon or contact Zalo OA Wind Official.</p><p>您的预约已确认。请准时到店。如需更改预约，请致电沙龙或联系 Zalo OA Wind Official。</p><p>'+SALON_ADDR+'</p></div>';
  MailApp.sendEmail({to:d.email,subject:'Xác nhận lịch hẹn tại '+SALON_NAME,body:detail,htmlBody:html,name:SALON_NAME});
}
