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
  if(sh.getLastRow()===0) {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
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
  var all=sh.getRange(2,1,sh.getLastRow()-1,HEADERS.length).getValues();
  return all.filter(function(r){return r[1] && !/^(đã hủy|đã huỷ|hủy|huỷ|cancelled|canceled|không đến|no.show)$/i.test(String(r[8]).trim());}).map(function(r){
    var begin=start_(dateKey_(r[2]),timeKey_(r[3]));
    if(!Number.isFinite(begin)) throw new Error('legacy_date_invalid');
    var duration=Number(r[10])||duration_(String(r[1]));
    if(duration<=0||duration>1440)throw new Error('invalid_duration');
    return {start:begin,end:begin+duration*60000,staff:staffName_(String(r[9])),id:String(r[12]||''),sourceRow:all.indexOf(r)+2,units:r[8]==='Khóa giờ'?CAPACITY:1,row:r};
  });
}
function peak_(rows,staff,a,b) {
  var events=[];
  rows.forEach(function(r){
    // Old unassigned bookings reserve capacity across every stylist until assigned.
    if((!r.staff || r.staff===staff) && r.start<b && r.end>a) {
      events.push([Math.max(a,r.start),r.units||1]); events.push([Math.min(b,r.end),-(r.units||1)]);
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
    return i>=0 && loads[i]<CAPACITY ? STAFF[i] : '';
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
function supported_(s){return typeof s==='string' && Object.prototype.hasOwnProperty.call(DURATIONS,s);}
function validStaff_(s){return !s||s==='auto'||STAFF.indexOf(s)>=0;}
function doGet(e) {
  try {
    var p=(e&&e.parameter)||{};
    if(p.action!=='availability')return json_({ok:true,version:2,msg:'WIND booking endpoint đang chạy.'});
    if(!supported_(p.service) || !validStaff_(p.stylist) || !/^\d{4}-\d{2}-\d{2}$/.test(p.date||''))return json_({ok:false,error:'invalid_service_or_date'});
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
    if(typeof d.name!=='string'||d.name.trim().length<2||d.name.length>150||typeof d.phone!=='string'||!supported_(d.service)||!validStaff_(d.stylist)||!validSlot_(String(d.date),String(d.time),Date.now()))return json_({ok:false,error:'invalid_fields'});
    var phone=String(d.phone).replace(/[^0-9]/g,'');
    if(!/^\d{8,15}$/.test(phone))return json_({ok:false,error:'invalid_phone'});
    if(!/^[a-zA-Z0-9-]{16,80}$/.test(d.requestId||''))return json_({ok:false,error:'invalid_request_id'});
    if(d.email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email))return json_({ok:false,error:'invalid_email'});
    if(!lock.tryLock(25000))return json_({ok:false,error:'busy'});
    var sh=getSheet_(),rows=rows_(sh);
    var duplicate=rows.filter(function(r){return r.id===d.requestId;})[0];
    if(duplicate)return json_({ok:true,id:duplicate.id,stylist:duplicate.staff,date:dateKey_(duplicate.row[2]),time:timeKey_(duplicate.row[3]),service:String(duplicate.row[1])});
    var a=start_(d.date,d.time),b=a+DURATIONS[d.service]*60000;
    // A network retry with a new browser request ID must not create a duplicate.
    var same=rows.filter(function(r){return r.start===a&&String(r.row[5]).replace(/\D/g,'')===phone&&String(r.row[1])===d.service;})[0];
    if(same)return json_({ok:true,id:same.id||'existing',stylist:same.staff,date:d.date,time:d.time,service:d.service});
    assigned=assign_(rows,a,b,d.stylist||'auto',d.foreign===true);
    if(!assigned)return json_({ok:false,error:'slot_full'});
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
  try{MailApp.sendEmail(NOTIFY_EMAIL,'Lịch hẹn đã xác nhận — '+d.name,detail+'\n\n'+SpreadsheetApp.getActiveSpreadsheet().getUrl());}catch(err){console.error('Salon email failed; appointment remains saved');}
  if(!d.email)return;
  var html=confirmationHtml_(d,id,stylist);
  MailApp.sendEmail({to:d.email,subject:'Xác nhận lịch hẹn tại '+SALON_NAME,body:detail,htmlBody:html,name:SALON_NAME});
}


function confirmationHtml_(d,id,stylist) {
  var font="font-family:Montserrat,Arial,Helvetica,sans-serif;";
  function row(label,en,value) {
    return '<tr><th scope="row" width="40%" style="'+font+'text-align:left;vertical-align:top;padding:14px 12px;border:1px solid #dddddd;font-size:12px;font-weight:600;color:#111111">'+label+'<br><span style="font-size:10px;font-weight:400;color:#777777">'+en+'</span></th><td style="'+font+'padding:14px 12px;border:1px solid #dddddd;font-size:13px;line-height:1.6;word-break:break-word;color:#111111">'+esc_(value)+'</td></tr>';
  }
  var date=String(d.date||'').split('-').reverse().join('/');
  var rows=row('Mã lịch hẹn','Booking reference',id)
    +row('Khách hàng','Guest name',d.name)
    +row('Số điện thoại','Phone number',d.phone)
    +row('Dịch vụ','Service',d.service)
    +row('Ngày hẹn','Appointment date',date)
    +row('Giờ hẹn','Appointment time',d.time+' · Giờ Việt Nam (UTC+7)')
    +row('Nhà tạo mẫu','Stylist',stylist)
    +row('Địa điểm','Salon address',SALON_ADDR)
    +row('Lời nhắn','Guest notes',d.note||'—');
  return '<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap" rel="stylesheet"></head>'
    +'<body style="margin:0;padding:0;background:#f2f2f2;'+font+'"><div style="display:none;max-height:0;overflow:hidden">WIND đã xác nhận lịch hẹn của bạn: '+esc_(d.time)+' ngày '+esc_(date)+'.</div>'
    +'<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f2f2f2"><tr><td align="center" style="padding:24px 10px">'
    +'<table role="presentation" width="600" cellspacing="0" cellpadding="0" style="width:100%;max-width:600px;background:#ffffff;border:1px solid #dddddd">'
    +'<tr><td style="background:#111111;color:#ffffff;padding:30px 24px;'+font+'"><img src="https://github.com/daoduykhuyen2-beep/Windhairhouse/blob/main/wind-logo.png?raw=true" width="76" height="70" alt="WIND Hair House" style="display:block;width:76px;height:auto;border:0;margin:0 0 18px"><div style="font-size:28px;letter-spacing:6px;font-weight:700">WIND</div><div style="font-size:11px;letter-spacing:3px;margin-top:6px">HAIR HOUSE</div><div style="font-size:10px;color:#cccccc;margin-top:16px">Devoted to the Art of Hair.</div></td></tr>'
    +'<tr><td style="padding:28px 20px 14px;text-align:center;'+font+'"><div style="font-size:10px;letter-spacing:2px;font-weight:600">ĐÃ XÁC NHẬN · CONFIRMED</div><h1 style="font-size:21px;line-height:1.5;margin:12px 0 4px;font-weight:700">XÁC NHẬN LỊCH HẸN</h1><div style="font-size:12px;color:#777777">Appointment Confirmation</div></td></tr>'
    +'<tr><td style="padding:8px 20px 24px;'+font+'"><p style="font-size:13px;line-height:1.8">Chào '+esc_(d.name)+',<br>WIND đã xác nhận và sắp xếp nhà tạo mẫu cho bạn. Hẹn gặp bạn tại salon!</p>'
    +'<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;table-layout:fixed">'+rows+'</table>'
    +'<p style="font-size:12px;line-height:1.8;margin:20px 0">Vui lòng đến đúng giờ hẹn. Nếu cần đổi hoặc hủy lịch, hãy liên hệ WIND qua điện thoại hoặc Zalo OA.</p>'
    +'<table role="presentation" cellspacing="0" cellpadding="0"><tr><td style="background:#111111;padding:13px 16px"><a href="tel:0981866736" style="'+font+'font-size:11px;font-weight:600;color:#ffffff;text-decoration:none">GỌI WIND</a></td><td width="10"></td><td style="background:#111111;padding:13px 16px"><a href="https://zalo.me/4548028938819960082" style="'+font+'font-size:11px;font-weight:600;color:#ffffff;text-decoration:none">ZALO OA</a></td></tr></table>'
    +'<p lang="en" style="font-size:11px;line-height:1.8;color:#666666;margin:22px 0 12px">Your appointment is confirmed. Please arrive on time. To reschedule or cancel, call WIND or contact our Zalo OA.</p>'
    +'<p lang="zh" style="font-size:11px;line-height:1.8;color:#666666">您的预约已确认。请准时到店。如需更改或取消预约，请致电 WIND 或联系 Zalo OA。</p>'
    +'<div style="border-top:1px solid #dddddd;padding-top:18px;margin-top:22px;text-align:center;font-size:12px;font-weight:600;line-height:1.8">Cảm ơn bạn đã lựa chọn WIND HAIR HOUSE!<br><span style="font-size:10px;font-weight:400;color:#777777">Thank you for choosing WIND HAIR HOUSE.</span></div>'
    +'<p style="font-size:10px;color:#777777;line-height:1.8;margin:18px 0 0">Đây là xác nhận lịch hẹn, không phải xác nhận thanh toán. Chi phí dịch vụ được tư vấn và báo rõ trước khi thực hiện.</p></td></tr>'
    +'<tr><td style="background:#111111;color:#ffffff;padding:18px 20px;text-align:center;'+font+'font-size:10px;line-height:1.8">'+esc_(SALON_ADDR)+'<br>09:00–20:00 · '+esc_(SALON_PHONE)+'</td></tr></table></td></tr></table></body></html>';
}

// Private salon controls, available only inside the bound Google Sheet.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('WIND · Quản lý lịch').addItem('Mở bảng quản lý lịch','showWIND').addToUi();
}
function showWIND() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(WIND_ADMIN_HTML).setWidth(1000).setHeight(720),'WIND Hair House · Quản lý lịch hẹn');
}
function adminData_(){
  var sh=getSheet_();return {sheet:sh,values:sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,HEADERS.length).getValues():[]};
}
function version_(r){return JSON.stringify(r);}
function checkedRow_(sh,n,version){
  n=Number(n);if(!Number.isInteger(n)||n<2||n>sh.getLastRow())throw Error('Dòng lịch đã thay đổi. Vui lòng tải lại.');
  var r=sh.getRange(n,1,1,HEADERS.length).getValues()[0];
  if(version_(r)!==version)throw Error('Lịch vừa được người khác cập nhật. Vui lòng tải lại trước khi sửa.');return r;
}
function windAdminRead(date){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date||''))throw Error('Chọn ngày hợp lệ.');
  var data=adminData_(),day=start_(date,'00:00');
  var list=data.values.map(function(r,i){
    var a=start_(dateKey_(r[2]),timeKey_(r[3])),mins=Number(r[10])||duration_(String(r[1]));
    return {rowNo:i+2,version:version_(r),service:String(r[1]||''),date:dateKey_(r[2]),time:timeKey_(r[3]),name:String(r[4]||''),phone:String(r[5]||'').replace(/^'/,''),email:String(r[6]||''),note:String(r[7]||''),status:String(r[8]||'Mới'),stylist:staffName_(String(r[9])),duration:mins,end:Number.isFinite(a)?Utilities.formatDate(new Date(a+mins*60000),TZ,'dd/MM/yyyy HH:mm'):'Cần kiểm tra ngày/giờ',id:String(r[12]||''),foreign:r[13]==='Khách nước ngoài',start:a};
  }).filter(function(r){return r.service&&(r.date===date||(r.start<day&&r.start+r.duration*60000>day));}).sort(function(a,b){return a.start-b.start;});
  return {date:date,staff:STAFF,capacity:CAPACITY,rows:list};
}
function windAdminSave(d){
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(25000))throw Error('Hệ thống đang lưu lịch khác. Vui lòng thử lại.');
  try{
    var sh=getSheet_(),old=d.rowNo?checkedRow_(sh,d.rowNo,d.version):null;
    if(!old){
      if(!/^[a-zA-Z0-9-]{16,80}$/.test(d.requestId||''))throw Error('Vui lòng tải lại màn hình quản lý rồi lưu lịch.');
      var prior=adminData_().values.filter(function(r){return String(r[12])===d.requestId;})[0];
      if(prior)return {ok:true,id:String(prior[12]),stylist:String(prior[9]),existing:true};
    }
    var duration=Number(d.duration),a=start_(String(d.date),String(d.time));
    if(!d.name||String(d.name).trim().length<2||!d.service||!/^\d{8,15}$/.test(String(d.phone).replace(/\D/g,''))||!validStaff_(d.stylist)||!Number.isInteger(duration)||duration<15||duration>1440)throw Error('Kiểm tra tên, điện thoại, dịch vụ, stylist và thời lượng (15–1440 phút).');
    if(!validSlot_(String(d.date),String(d.time),Date.now())&&!(old&&a===start_(dateKey_(old[2]),timeKey_(old[3]))))throw Error('Giờ mới phải còn ở tương lai, từ 09:00 đến 20:00, cách nhau 30 phút, trong 180 ngày.');
    if(old&&old[8]==='Khóa giờ')throw Error('Hãy hủy dòng khóa giờ rồi tạo lại nếu cần thay đổi.');
    var rows=rows_(sh).filter(function(r){return !old||r.sourceRow!==Number(d.rowNo);});
    var assigned=assign_(rows,a,a+duration*60000,d.stylist||'auto',d.foreign===true);
    if(!assigned)throw Error('Stylist đã đủ khách hoặc chưa phù hợp điều kiện nhận khách. Chọn giờ/thợ khác.');
    var stamp=old?old[0]:Utilities.formatDate(new Date(),TZ,'dd/MM/yyyy HH:mm');
    var id=old&&old[12]?old[12]:(d.requestId||'manual-'+Utilities.getUuid());
    var row=[stamp,safe_(d.service),d.date,d.time,safe_(d.name),"'"+String(d.phone).replace(/\D/g,''),safe_(d.email),safe_(d.note),old&&!/^(Mới|Đã hủy|Đã huỷ|Không đến|cancelled|canceled)$/.test(String(old[8]))?old[8]:'Đã xác nhận',assigned,duration,Utilities.formatDate(new Date(a+duration*60000),TZ,'yyyy-MM-dd HH:mm'),id,d.foreign===true?'Khách nước ngoài':'Khách Việt Nam'];
    if(old)sh.getRange(Number(d.rowNo),1,1,HEADERS.length).setValues([row]);else sh.appendRow(row);
    SpreadsheetApp.flush();return {ok:true,id:id,stylist:assigned};
  }finally{lock.releaseLock();}
}
function windAdminStatus(n,version,status){
  var allowed=['Đã xác nhận','Đã đến','Đang làm','Hoàn tất','Đã hủy','Không đến'];
  if(allowed.indexOf(status)<0)throw Error('Trạng thái không hợp lệ.');
  var lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    var sh=getSheet_(),old=checkedRow_(sh,n,version);
    if(old[8]==='Khóa giờ'&&status!=='Đã hủy')throw Error('Chỉ có thể hủy dòng khóa giờ.');
    if(/^(Đã hủy|Không đến)$/.test(String(old[8]))&&!/^(Đã hủy|Không đến)$/.test(status))throw Error('Dùng nút Sửa để kiểm tra chỗ trống trước khi khôi phục lịch.');
    sh.getRange(Number(n),9).setValue(status);SpreadsheetApp.flush();return {ok:true};
  }finally{lock.releaseLock();}
}
function windAdminBlock(d){
  var a=start_(String(d.date),String(d.time)),duration=Number(d.duration);
  if(STAFF.indexOf(d.stylist)<0||!/^\d{4}-\d{2}-\d{2}$/.test(d.date||'')||!/^([01]\d|2[0-3]):[03]0$/.test(d.time||'')||!Number.isFinite(a)||a<=Date.now()||!Number.isInteger(duration)||duration<30||duration>1440)throw Error('Chọn thợ, ngày giờ tương lai và thời lượng khóa 30–1440 phút.');
  var lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    var sh=getSheet_(),b=a+duration*60000;
    if(rows_(sh).some(function(r){return (!r.staff||r.staff===d.stylist)&&r.start<b&&r.end>a;}))throw Error('Khoảng này đã có khách hoặc lịch chưa gán thợ. Vui lòng sắp xếp lại trước khi khóa.');
    sh.appendRow([Utilities.formatDate(new Date(),TZ,'dd/MM/yyyy HH:mm'),'Khóa giờ / Nghỉ',d.date,d.time,'KHÓA GIỜ','','',safe_(d.note),'Khóa giờ',d.stylist,duration,Utilities.formatDate(new Date(b),TZ,'yyyy-MM-dd HH:mm'),'block-'+Utilities.getUuid(),'Nội bộ']);
    SpreadsheetApp.flush();return {ok:true};
  }finally{lock.releaseLock();}
}

const WIND_ADMIN_HTML = "<!doctype html><html lang=\"vi\"><head><meta charset=\"utf-8\"><style>\n*{box-sizing:border-box}body{margin:0;padding:20px;background:#f7f6f2;color:#181818;font:13px Arial,sans-serif}h2{font-size:20px;margin:0 0 8px}p{line-height:1.6}.layout{display:grid;grid-template-columns:320px 1fr;gap:22px}.card{background:#fff;border:1px solid #ddd;padding:16px;border-radius:8px;margin-bottom:12px}label{display:block;font-size:12px;font-weight:bold;margin:10px 0 5px}input,select,textarea,button{font:inherit;padding:9px;border:1px solid #ccc;border-radius:5px;width:100%}textarea{min-height:55px}button{background:#181818;color:#fff;cursor:pointer;margin-top:8px}button.light{background:#fff;color:#181818}button:disabled{opacity:.5;cursor:wait}.pair{display:grid;grid-template-columns:1fr 1fr;gap:10px}.muted{color:#666;font-size:12px}#message{padding:12px;background:#eee;white-space:pre-line;line-height:1.5;margin:12px 0;min-height:40px}#message.error{background:#fff0ed;color:#9a260e}.rowhead{display:flex;justify-content:space-between;gap:8px}.rowhead strong{font-size:15px}.badge{font-size:11px;border-radius:10px;background:#eee;padding:4px 7px;white-space:nowrap}.warning{border-left:4px solid #d28b1f}details{line-height:1.6;margin:12px 0}summary{cursor:pointer;font-weight:bold}.actions{display:flex;gap:6px;flex-wrap:wrap}.actions button{width:auto;padding:6px 9px;font-size:12px}.hidden{display:none!important}.footer{font-size:11px;color:#666;line-height:1.6}#count{margin:10px 0;font-weight:bold}\n</style></head><body>\n<h2>WIND · Quản lý lịch hẹn</h2><div class=\"muted\">3 khách / stylist · Nhận lịch 09:00–20:00 · Giờ Việt Nam · Có thể kết thúc qua đêm</div>\n<div id=\"message\" role=\"status\">Đang tải…</div><div class=\"layout\"><section>\n<div class=\"card\"><label for=\"viewDate\">XEM LỊCH NGÀY</label><div class=\"pair\"><input id=\"viewDate\" type=\"date\"><button class=\"light\" id=\"refresh\" type=\"button\">Tải lại lịch</button></div></div>\n<div class=\"card\"><h3 id=\"formTitle\">Thêm lịch điện thoại / Zalo</h3><form id=\"bookingAdmin\">\n<label for=\"guest\">Tên khách</label><input id=\"guest\" required maxlength=\"150\">\n<label for=\"phone\">Số điện thoại</label><input id=\"phone\" type=\"tel\" required>\n<label for=\"service\">Dịch vụ</label><input id=\"service\" list=\"services\" required><datalist id=\"services\"><option>Cắt tóc</option><option>Uốn tóc</option><option>Duỗi tóc</option><option>Nhuộm tóc</option><option>Tẩy + Nhuộm</option><option>Nối tóc</option></datalist>\n<div class=\"pair\"><div><label for=\"bookDate\">Ngày hẹn</label><input id=\"bookDate\" type=\"date\" required></div><div><label for=\"time\">Giờ bắt đầu</label><select id=\"time\" required></select></div></div>\n<div class=\"pair\"><div><label for=\"duration\">Thời lượng (phút)</label><input id=\"duration\" type=\"number\" min=\"15\" max=\"1440\" step=\"1\" value=\"60\" required></div><div><label for=\"foreign\">Nhóm khách</label><select id=\"foreign\"><option value=\"false\">Khách Việt Nam</option><option value=\"true\">Khách nước ngoài</option></select></div></div>\n<label for=\"staff\">Stylist</label><select id=\"staff\"><option value=\"auto\">Salon sắp xếp</option><option>Phan Cường</option><option>Trung Ngọc</option><option>Đào Duy Khuyến</option></select>\n<label for=\"email\">Email (nếu có)</label><input id=\"email\" type=\"email\"><label for=\"note\">Ghi chú / Nguồn lịch</label><textarea id=\"note\" placeholder=\"Điện thoại / Zalo / Trực tiếp…\"></textarea>\n<button type=\"submit\">KIỂM TRA CHỖ VÀ LƯU</button><button id=\"clear\" class=\"light\" type=\"button\">Tạo lịch mới / Bỏ chỉnh sửa</button>\n<p class=\"footer\">Lưu ở đây không tự gửi email. Nhân viên báo lại lịch cho khách qua kênh đang liên hệ. Khách có thể chọn Khuyến trực tiếp khi còn chỗ; salon tự sắp xếp ưu tiên Cường và Trung cho khách Việt.</p></form></div>\n<details class=\"card\"><summary>Khóa giờ thợ / Nghỉ</summary><form id=\"blockAdmin\"><label for=\"blockStaff\">Stylist</label><select id=\"blockStaff\"><option>Phan Cường</option><option>Trung Ngọc</option><option>Đào Duy Khuyến</option></select><label for=\"blockDate\">Ngày</label><input id=\"blockDate\" type=\"date\" required><label for=\"blockTime\">Bắt đầu</label><input id=\"blockTime\" type=\"time\" step=\"1800\" value=\"09:00\" required><label for=\"blockDuration\">Khóa bao nhiêu phút?</label><input id=\"blockDuration\" type=\"number\" min=\"30\" max=\"1440\" value=\"660\" required><label for=\"blockNote\">Lý do</label><input id=\"blockNote\" placeholder=\"Nghỉ / đào tạo / việc riêng…\"><button>KIỂM TRA VÀ KHÓA GIỜ</button><p class=\"footer\">Khóa toàn bộ 3 chỗ của thợ trong khoảng đã chọn. Nếu có khách, cần sắp xếp lại trước. Để mở lại, hủy dòng khóa giờ.</p></form></details>\n</section><section><div id=\"count\"></div><div id=\"appointments\"></div>\n<details class=\"card\" open><summary>Quy trình sử dụng</summary><ol><li>Lịch trên website tự xuất hiện tại đây.</li><li>Lịch điện thoại/Zalo: nhập form bên trái rồi báo lại khách.</li><li>Đổi lịch: bấm Sửa, chọn ngày/giờ/thợ mới, kiểm tra và lưu. Không tạo thêm dòng cho cùng lịch.</li><li>Khách hủy hoặc không đến: cập nhật trạng thái để trả chỗ.</li><li>Dòng chưa có stylist phải được kiểm tra và gán đúng người. Hệ thống đang tính thận trọng vào cả ba thợ.</li><li>Giữ thời lượng tối đa: cắt 60, uốn/duỗi 240, nhuộm 90, tẩy + nhuộm 360, nối 240 phút. Dịch vụ khác nhập thời lượng đã tư vấn.</li><li>Nếu vừa có người khác sửa lịch, tải lại để tránh ghi đè.</li></ol></details>\n<p class=\"footer\">Dữ liệu khách chỉ hiển thị cho người có quyền truy cập bảng tính. Trang website công khai chỉ nhận danh sách giờ trống.</p></section></div>\n<script>\nconst $=id=>document.getElementById(id);let editing=null,rows=[],busy=false,manualRequestId='',manualFingerprint='';\nfunction msg(s,error=false){$('message').textContent=s;$('message').className=error?'error':'';}\nfunction setBusy(on){busy=on;document.querySelectorAll('button').forEach(b=>b.disabled=on);}\nfunction call(name,args,done){setBusy(true);google.script.run.withSuccessHandler(r=>{setBusy(false);done(r);}).withFailureHandler(e=>{setBusy(false);msg(e.message||String(e),true);})[name](...args);}\nfunction text(tag,value,parent,cls){const el=document.createElement(tag);el.textContent=value;if(cls)el.className=cls;parent.append(el);return el;}\nfunction action(parent,label,fn){const b=text('button',label,parent,'light');b.type='button';b.onclick=fn;return b;}\nfunction load(){msg('Đang tải lịch…');call('windAdminRead',[$('viewDate').value],r=>{rows=r.rows;render();msg('Đã tải '+rows.length+' lịch. Thay đổi được kiểm tra lại trước khi lưu.');});}\nfunction render(){\n $('appointments').replaceChildren();$('count').textContent=$('viewDate').value+' · '+rows.length+' dòng lịch';\n if(!rows.length)text('p','Chưa có lịch trong ngày này.',$('appointments'),'card muted');\n rows.forEach(r=>{const card=text('article','',$('appointments'),'card'+(!r.stylist?' warning':''));const head=text('div','',card,'rowhead');text('strong',r.time+' · '+r.name,head);text('span',r.status,head,'badge');text('p',r.service+' · '+r.duration+' phút\\n'+(r.stylist||'⚠ Chưa gán thợ')+'\\nKết thúc: '+r.end,card).style.whiteSpace='pre-line';if(r.phone)text('p',r.phone+(r.email?' · '+r.email:''),card,'muted');if(r.note)text('p',r.note,card,'muted');const actions=text('div','',card,'actions');\n   if(r.status!=='Khóa giờ')action(actions,'Sửa / Đổi lịch',()=>edit(r));\n   if(r.status!=='Khóa giờ'&&!['Đã hủy','Không đến'].includes(r.status))['Đã đến','Đang làm','Hoàn tất','Không đến'].forEach(s=>action(actions,s,()=>status(r,s)));\n   if(r.status!=='Đã hủy'){let armed=false;const cancelButton=action(actions,r.status==='Khóa giờ'?'Mở lại giờ':'Hủy lịch',()=>{if(!armed){armed=true;cancelButton.textContent=r.status==='Khóa giờ'?'XÁC NHẬN MỞ GIỜ':'XÁC NHẬN HỦY';msg('Kiểm tra đúng lịch của '+r.name+' rồi bấm nút xác nhận lần nữa. Bấm Tải lại lịch nếu muốn bỏ thao tác.');return;}status(r,'Đã hủy');});}\n });\n}\nfunction status(r,s){call('windAdminStatus',[r.rowNo,r.version,s],()=>load());}\nfunction clear(){editing=null;manualRequestId='';manualFingerprint='';$('bookingAdmin').reset();$('bookDate').value=$('viewDate').value;$('formTitle').textContent='Thêm lịch điện thoại / Zalo';}\nfunction edit(r){editing=r;$('formTitle').textContent='Sửa lịch · dòng '+r.rowNo;for(const [id,key] of Object.entries({guest:'name',phone:'phone',service:'service',bookDate:'date',time:'time',duration:'duration',email:'email',note:'note'}))$(id).value=r[key];$('foreign').value=String(r.foreign);$('staff').value=r.stylist||'auto';$('guest').focus();msg('Đang sửa lịch dòng '+r.rowNo+'. Hệ thống sẽ kiểm tra chỗ trống khi lưu.');}\n$('bookingAdmin').onsubmit=e=>{e.preventDefault();if(busy)return;const d={name:$('guest').value.trim(),phone:$('phone').value,service:$('service').value.trim(),date:$('bookDate').value,time:$('time').value,duration:Number($('duration').value),foreign:$('foreign').value==='true',stylist:$('staff').value,email:$('email').value,note:$('note').value};if(editing){d.rowNo=editing.rowNo;d.version=editing.version;}else{const fingerprint=JSON.stringify(d);if(fingerprint!==manualFingerprint){manualRequestId='manual-'+(typeof crypto.randomUUID==='function'?crypto.randomUUID():Date.now()+'-'+Math.random().toString(36).slice(2));manualFingerprint=fingerprint;}d.requestId=manualRequestId;}call('windAdminSave',[d],r=>{clear();load();msg('Đã lưu lịch cho '+r.stylist+'. Vui lòng báo lại khách.');});};\n$('blockAdmin').onsubmit=e=>{e.preventDefault();if(busy)return;call('windAdminBlock',[{date:$('blockDate').value,time:$('blockTime').value,duration:Number($('blockDuration').value),stylist:$('blockStaff').value,note:$('blockNote').value}],()=>load());};\n$('service').onchange=()=>{const d={'Cắt tóc':60,'Uốn tóc':240,'Duỗi tóc':240,'Nhuộm tóc':90,'Tẩy + Nhuộm':360,'Nối tóc':240};if(d[$('service').value])$('duration').value=d[$('service').value];};\n$('clear').onclick=clear;$('refresh').onclick=load;$('viewDate').onchange=()=>{clear();$('blockDate').value=$('viewDate').value;load();};\nfor(let m=540;m<=1200;m+=30){const t=String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0');$('time').add(new Option(t,t));}\nconst today=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Ho_Chi_Minh',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());$('viewDate').value=today;$('bookDate').value=today;$('blockDate').value=today;load();\n</script></body></html>\n";

