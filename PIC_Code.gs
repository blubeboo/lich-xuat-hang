/**
 * TIẾN ĐỘ NHÂN VIÊN THEO NGÀY - Backend API (v1)
 * Đọc file "Order Processing Tracking", sheet "3. Detail Action by PIC".
 *
 * Đây là project Apps Script RIÊNG (script.google.com → Dự án mới),
 * không đụng tới Gantt_Code.gs hay OrderProcessing.gs của file Sheet.
 *
 * Cột dùng:
 *   C Order Number | D Carton No. | E Product Type | F Process stage
 *   G Time Plan    | I Working date | J PIC | R Status (% của dòng)
 *
 * Trả về các dòng có Working date trong khoảng hôm nay −30 … +14 ngày.
 * Trang web tự gom theo ngày → theo người → theo đơn + tiến trình.
 *
 * CÀI ĐẶT 1 LẦN: chọn hàm setupTrigger → Run (cấp quyền nếu được hỏi).
 * Gọi kèm ?fresh=1 để tính lại ngay (nút Làm mới trên trang).
 */

var SPREADSHEET_ID = '17n7azm6uxKuYIjkAJrZkCFXu4XWH_BTvyEMrVnsBjvI';
var SHEET_NAME = '3. Detail Action by PIC';
var DAYS_BACK = 30;
var DAYS_AHEAD = 14;

var CACHE_PREFIX = 'pic_v1_';
var CACHE_TTL = 21600;   // 6 tiếng
var CHUNK = 90000;

var STAGE_LABEL = { '1 Prepare time': 'Lấy', '2 Stamp time': 'Dán', '3 Check-Pack Time': 'Kiểm' };

// ================== API ==================
function doGet(e) {
  try {
    var fresh = e && e.parameter && e.parameter.fresh;
    var json = fresh ? null : readCache();
    if (!json) json = refreshCache();
    return out(json, e);
  } catch (err) {
    return out(JSON.stringify({ ok: false, error: String(err) }), e);
  }
}

function refreshCache() {
  var json = JSON.stringify({ ok: true, updated: new Date().toISOString(), rows: buildRows() });
  var cache = CacheService.getScriptCache();
  var parts = {}, n = Math.ceil(json.length / CHUNK);
  for (var i = 0; i < n; i++) parts[CACHE_PREFIX + i] = json.substr(i * CHUNK, CHUNK);
  parts[CACHE_PREFIX + 'n'] = String(n);
  cache.putAll(parts, CACHE_TTL);
  return json;
}

function readCache() {
  var cache = CacheService.getScriptCache();
  var n = cache.get(CACHE_PREFIX + 'n');
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < +n; i++) keys.push(CACHE_PREFIX + i);
  var got = cache.getAll(keys), s = '';
  for (var k = 0; k < keys.length; k++) { if (got[keys[k]] == null) return null; s += got[keys[k]]; }
  return s;
}

// Chạy 1 lần để bật tính sẵn mỗi 5 phút
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshCache').timeBased().everyMinutes(5).create();
  refreshCache();
  Logger.log('Đã bật tính sẵn mỗi 5 phút.');
}

// Kiểm tra nhanh: Run → xem Nhật ký thực thi
function testPIC() {
  var rows = buildRows();
  var today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  var t = rows.filter(function (r) { return r.d === today; });
  Logger.log('Tổng dòng trong khoảng: ' + rows.length + ' · Dòng hôm nay: ' + t.length);
  Logger.log(JSON.stringify(t.slice(0, 10)));
}

// ================== ĐỌC DỮ LIỆU ==================
function buildRows() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Không tìm thấy sheet "' + SHEET_NAME + '"');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var cj = sh.getRange(1, 3, lastRow, 8).getValues();   // C..J
  var rr = sh.getRange(1, 18, lastRow, 1).getValues();  // R

  // Tìm dòng tiêu đề ("Order Number" ở cột C)
  var start = 4;
  for (var h = 0; h < Math.min(cj.length, 15); h++) {
    if (String(cj[h][0]).trim().toLowerCase() === 'order number') { start = h + 1; break; }
  }

  var zone = tz();
  var now = new Date();
  var from = new Date(now.getTime() - DAYS_BACK * 86400000);
  var to = new Date(now.getTime() + DAYS_AHEAD * 86400000);
  var fromK = Utilities.formatDate(from, zone, 'yyyy-MM-dd');
  var toK = Utilities.formatDate(to, zone, 'yyyy-MM-dd');

  var rows = [];
  for (var i = start; i < cj.length; i++) {
    var r = cj[i];
    var code = String(r[0] || '').trim();      // C
    var stage = String(r[3] || '').trim();     // F
    if (!code || !stage) continue;
    var d = dayKey(r[6], zone);                // I
    if (!d || d < fromK || d > toK) continue;

    rows.push({
      d: d,
      p: String(r[7] || '').trim(),            // J PIC
      c: code,
      n: cartonStr(r[1], zone),                // D Carton
      t: String(r[2] || '').trim(),            // E Product Type
      s: STAGE_LABEL[stage] || stage,          // F
      h: num(r[4]),                            // G Time Plan
      r: pct(rr[i][0])                         // R Status (null = trống)
    });
  }
  return rows;
}

// ================== TIỆN ÍCH ==================
function tz() { return Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh'; }

function dayKey(v, zone) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, zone, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);          // MM/dd/yyyy
  if (m) return m[3] + '-' + pad(m[1]) + '-' + pad(m[2]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  return '';
}

function cartonStr(v, zone) {
  if (v instanceof Date) return Utilities.formatDate(v, zone, 'd/M');
  return String(v == null ? '' : v).trim();
}

function pad(x) { return ('0' + x).slice(-2); }

function num(v) {
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

// Status: 1 → 100, 0.5 → 50, "50%" → 50, trống → null
function pct(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return Math.round((v <= 1 ? v * 100 : v) * 10) / 10;
  var s = String(v).trim();
  if (!s) return null;
  var n = parseFloat(s.replace('%', '').replace(',', '.'));
  if (isNaN(n)) return null;
  if (s.indexOf('%') < 0 && n <= 1) n = n * 100;
  return Math.round(n * 10) / 10;
}

function out(json, e) {
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
