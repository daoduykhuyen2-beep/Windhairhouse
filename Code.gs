/**
 * WIND HAIR HOUSE — Nhận đặt lịch từ website → Google Sheet + Email xác nhận
 * ------------------------------------------------------------------------
 * Cách hoạt động: Form trên website windhairhouse.vn gửi dữ liệu tới đây (doPost).
 *   1) Thêm 1 dòng vào Google Sheet
 *   2) Gửi email báo cho salon (có khách mới)
 *   3) Gửi email xác nhận cho khách (nếu khách có điền email)
 *
 * KHÔNG cần sửa gì để chạy — chỉ cần chỉnh 5 dòng CẤU HÌNH bên dưới nếu muốn.
 */

// ===================== CẤU HÌNH =====================
const SHEET_NAME   = 'Lịch hẹn';                 // tên tab trong Google Sheet
const NOTIFY_EMAIL = 'daoduykhuyen2@gmail.com';  // email nhận báo có khách mới
const SALON_NAME   = 'WIND HAIR HOUSE';
const SALON_PHONE  = '098.186.6736';
const SALON_ADDR   = 'Tòa S8.03, Vinhomes Grand Park, Thủ Đức, TP.HCM';
const FORM_SECRET  = 'wind_d375b0682cd04279'; // mã bí mật — PHẢI trùng mã trong index.html
const MAX_LEN      = 500;                      // giới hạn ký tự mỗi ô
// ====================================================

const HEADERS = ['Thời gian gửi', 'Dịch vụ', 'Ngày hẹn', 'Khung giờ',
                 'Họ tên', 'Số điện thoại', 'Email', 'Lời nhắn', 'Trạng thái'];

/** Nhận dữ liệu form gửi lên */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); }
      catch (err) { data = e.parameter || {}; }
    } else {
      data = (e && e.parameter) || {};
    }

    // ---- BẢO MẬT ----
    // 1) Bẫy bot: ô ẩn 'website' có dữ liệu => bot => bỏ qua (giả vờ thành công)
    if (data.website) return json_({ ok: true });
    // 2) Mã bí mật phải khớp
    if (data.token !== FORM_SECRET) return json_({ ok: false, error: 'unauthorized' });
    // 3) Bắt buộc đủ trường
    if (!data.name || !data.phone || !data.service || !data.date || !data.time)
      return json_({ ok: false, error: 'missing_fields' });
    // 4) Định dạng SĐT Việt Nam (0xxxxxxxxx hoặc 84xxxxxxxxx)
    var phone = String(data.phone).replace(/[^0-9]/g, '');
    if (!/^0\d{9}$/.test(phone) && !/^84\d{9}$/.test(phone))
      return json_({ ok: false, error: 'invalid_phone' });
    // 5) Cắt bớt độ dài mỗi ô, chặn payload rác
    ['name','service','date','time','email','note'].forEach(function (k) {
      if (data[k]) data[k] = String(data[k]).slice(0, MAX_LEN);
    });
    // 6) Chống gửi trùng: cùng 1 SĐT trong 60 giây chỉ nhận 1 lần
    var cache = CacheService.getScriptCache();
    if (cache.get('rl_' + phone)) return json_({ ok: false, error: 'too_soon' });
    cache.put('rl_' + phone, '1', 60);
    // ---- HẾT BẢO MẬT ----

    var sheet = getSheet_();
    var stamp = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');

    sheet.appendRow([
      stamp,
      data.service || '',
      data.date    || '',
      data.time    || '',
      data.name    || '',
      normalizePhone_(data.phone || ''),
      data.email   || '',
      data.note    || '',
      'Mới'
    ]);

    notifySalon_(data, stamp);
    if (data.email && /\S+@\S+\.\S+/.test(data.email)) confirmCustomer_(data);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Mở link Web App bằng trình duyệt sẽ chạy hàm này — dùng để kiểm tra */
function doGet(e) {
  getSheet_(); // đảm bảo tab + dòng tiêu đề đã tạo
  return json_({ ok: true, msg: 'WIND booking endpoint đang chạy.' });
}

/** Lấy (hoặc tạo) tab và dòng tiêu đề */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(8, 260); // cột Lời nhắn rộng hơn
  }
  return sheet;
}

/** Giữ số 0 đầu của SĐT (thêm dấu ' để Sheet coi là chữ, không cắt số 0) */
function normalizePhone_(p) {
  p = String(p).replace(/[^0-9+]/g, '');
  return p ? ("'" + p) : '';
}

/** Email báo cho salon */
function notifySalon_(d, stamp) {
  var subject = '🔔 Khách đặt lịch mới — ' + (d.name || 'Không tên') + ' (' + (d.phone || '') + ')';
  var body =
    'Có khách vừa đặt lịch trên website:\n\n' +
    'Họ tên   : ' + (d.name || '') + '\n' +
    'SĐT      : ' + (d.phone || '') + '\n' +
    'Dịch vụ  : ' + (d.service || '') + '\n' +
    'Ngày     : ' + (d.date || '') + '   Giờ: ' + (d.time || '') + '\n' +
    'Email    : ' + (d.email || '—') + '\n' +
    'Lời nhắn : ' + (d.note || '—') + '\n' +
    'Lúc gửi  : ' + stamp + '\n\n' +
    'Mở Google Sheet để xem toàn bộ lịch hẹn.';
  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

/** Email xác nhận gửi cho khách */
function confirmCustomer_(d) {
  var subject = 'Xác nhận lịch hẹn tại ' + SALON_NAME;
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:auto;color:#111">' +
      '<h2 style="letter-spacing:3px;margin:0 0 4px">' + SALON_NAME + '</h2>' +
      '<p style="color:#888;margin:0 0 18px;font-style:italic">Devoted to the Art of Hair</p>' +
      '<p>Chào ' + esc_(d.name || 'bạn') + ',</p>' +
      '<p>WIND đã nhận được yêu cầu đặt lịch của bạn:</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
        row_('Dịch vụ', d.service) +
        row_('Ngày hẹn', d.date) +
        row_('Khung giờ', d.time) +
        row_('Số điện thoại', d.phone) +
        (d.note ? row_('Lời nhắn', d.note) : '') +
      '</table>' +
      '<p style="margin-top:16px">Salon sẽ gọi lại để xác nhận chính thức. Cần đổi lịch, bạn gọi/Zalo <b>' + SALON_PHONE + '</b> nhé.</p>' +
      '<p style="color:#999;font-size:12px;margin-top:22px;border-top:1px solid #eee;padding-top:12px">' +
        SALON_NAME + ' · ' + SALON_ADDR + '<br>ĐT/Zalo: ' + SALON_PHONE +
      '</p>' +
    '</div>';
  MailApp.sendEmail({ to: d.email, subject: subject, htmlBody: html, name: SALON_NAME });
}

function row_(k, v) {
  return '<tr>' +
    '<td style="padding:7px 10px;border:1px solid #eee;color:#888;width:120px">' + k + '</td>' +
    '<td style="padding:7px 10px;border:1px solid #eee"><b>' + esc_(v || '') + '</b></td>' +
  '</tr>';
}
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
         .setMimeType(ContentService.MimeType.JSON);
}
