/**
 * SỔ ĐƠN SHOPEE — backend chạy trên Google Apps Script, gắn với 1 Google Sheet.
 * - Google Sheet là cơ sở dữ liệu (sheet "Đơn hàng").
 * - Ảnh chụp màn hình chỉ được chuyển tiếp tới Gemini để đọc, KHÔNG lưu ở đâu cả.
 *
 * Cần đặt 2 giá trị trong Project Settings > Script Properties:
 *   APP_PIN         mã PIN tự đặt, app trên điện thoại phải nhập đúng mới dùng được
 *   GEMINI_API_KEY  API key miễn phí lấy tại https://aistudio.google.com/app/apikey
 * Tuỳ chọn:
 *   GEMINI_MODEL    tên model, mặc định "gemini-flash-latest"
 */

const SCRIPT_VERSION = '2026-11-18c';   // app dùng số này để biết Apps Script có phải bản mới nhất
const SHEET_NAME = 'Đơn hàng';
const SHOP_SHEET = 'Cửa hàng';          // tên shop → link cửa hàng (dùng chung cho mọi đơn của shop)
const HEADERS = [
  'Mã đơn', 'Ngày đặt', 'Shop', 'Tên sản phẩm', 'Phân loại', 'SL', 'Giá mua (đ/món)',
  'Shopee Voucher (đ/đơn)', 'Voucher Shop (đ/đơn)', 'Shopee Xu (đ/đơn)',
  'Số món/đơn [Auto]', 'Voucher chia đều/món [Auto]', 'Xu chia đều/món [Auto]',
  'Thành tiền thực trả [Auto]', 'Đơn giá thực tế sau giảm [Auto]', 'Thời điểm nhập'
];

function doGet() {
  return out_({ ok: true, message: 'Sổ đơn Shopee đang chạy. Hãy dán địa chỉ này vào mục Cài đặt của app.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    const pin = props.getProperty('APP_PIN');
    if (!pin) return out_({ ok: false, error: 'Chưa đặt APP_PIN trong Script Properties.' });
    if (String(req.pin || '') !== pin) return out_({ ok: false, code: 'PIN', error: 'Sai mã PIN.' });

    switch (req.action) {
      case 'ping': return out_({ ok: true, version: SCRIPT_VERSION });
      case 'diag': return out_(diag_(props));
      case 'list': return out_({ ok: true, rows: listRows_(), shops: listShops_() });
      case 'setShopLink': return out_(setShopLink_(req.shop, req.link));
      case 'extract': return out_({ ok: true, order: extract_(req.images || [], props) });
      case 'save': return out_(save_(req.order, !!req.replace, req.originalId));
      case 'delete': return out_(remove_(req.id));
      default: return out_({ ok: false, error: 'Lệnh không hợp lệ.' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Google Sheet ---------------- */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e8eef0');
    sh.setFrozenRows(1);
    sh.getRange('A:E').setNumberFormat('@');          // mã đơn, ngày, chữ: giữ nguyên dạng văn bản
    sh.getRange('F:O').setNumberFormat('#,##0');
    sh.getRange('K1:O1').setBackground('#d9d9d9');
  }
  return sh;
}

function listRows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const txt = sh.getRange(2, 1, last - 1, 5).getDisplayValues();
  const num = sh.getRange(2, 6, last - 1, 5).getValues();
  const rows = [];
  for (let i = 0; i < txt.length; i++) {
    if (!txt[i][0]) continue;
    rows.push({
      maDon: txt[i][0], ngay: txt[i][1], shop: txt[i][2], ten: txt[i][3], phanLoai: txt[i][4],
      sl: Number(num[i][0]) || 0, gia: Number(num[i][1]) || 0,
      sv: Number(num[i][2]) || 0, vs: Number(num[i][3]) || 0, xu: Number(num[i][4]) || 0
    });
  }
  return rows;
}

/* ---------------- Link cửa hàng ---------------- */

function listShops_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHOP_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) {
    const name = String(r[0]).trim(), link = String(r[1]).trim();
    if (name && link) out[name] = link;
  });
  return out;
}

function setShopLink_(shop, link) {
  shop = String(shop || '').trim();
  link = String(link || '').trim();
  if (!shop) return { ok: false, error: 'Thiếu tên shop.' };
  if (link && !/^https?:\/\/\S+$/i.test(link)) return { ok: false, error: 'Link phải bắt đầu bằng http:// hoặc https://' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHOP_SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHOP_SHEET);
      sh.getRange(1, 1, 1, 2).setValues([['Shop', 'Link']]).setFontWeight('bold').setBackground('#e8eef0');
      sh.setFrozenRows(1);
      sh.getRange('A:B').setNumberFormat('@');
      sh.setColumnWidth(1, 240);
      sh.setColumnWidth(2, 460);
    }
    const last = sh.getLastRow();
    const names = last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); });
    const i = names.indexOf(shop);
    if (i !== -1 && link) sh.getRange(i + 2, 2).setValue(link);
    else if (i !== -1) sh.deleteRow(i + 2);
    else if (link) sh.appendRow([shop, link]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function deleteById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  let n = 0;
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === id) { sh.deleteRow(i + 2); n++; }
  }
  return n;
}

function int_(v) {
  // Số thật thì làm tròn; chuỗi như "150.000đ" / "150,000" thì bỏ hết ký tự không phải số
  // (dấu chấm ở VN là phân cách hàng nghìn, không phải dấu thập phân).
  const n = typeof v === 'number' ? Math.round(v) : parseInt(String(v == null ? '' : v).replace(/[^\d]/g, ''), 10);
  return isFinite(n) && n > 0 ? n : 0;
}

function save_(order, replace, originalId) {
  if (!order || !order.items || !order.items.length) return { ok: false, error: 'Đơn chưa có sản phẩm nào.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const tz = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    const id = String(order.maDon || '').trim() || ('DH-' + Utilities.formatDate(new Date(), tz, 'yyMMdd-HHmmss'));
    originalId = String(originalId || '').trim();

    if (id !== originalId) {
      const last = sh.getLastRow();
      const ids = last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
      if (ids.indexOf(id) !== -1 && !replace) {
        return { ok: false, code: 'DUP', error: 'Mã đơn ' + id + ' đã có trong sổ.' };
      }
    }
    if (originalId) deleteById_(sh, originalId);
    if (replace) deleteById_(sh, id);

    const sv = int_(order.sv), vs = int_(order.vs), xu = int_(order.xu);
    const start = sh.getLastRow() + 1;
    const values = [], formulas = [];
    order.items.forEach(function (it, i) {
      const r = start + i;
      values.push([id, String(order.ngay || ''), String(order.shop || ''), String(it.ten || ''),
        String(it.phanLoai || ''), int_(it.sl) || 1, int_(it.gia), sv, vs, xu]);
      formulas.push([
        '=COUNTIF($A$2:$A,A' + r + ')',
        '=(H' + r + '+I' + r + ')/K' + r,
        '=J' + r + '/K' + r,
        '=F' + r + '*G' + r + '-L' + r + '-M' + r,
        '=IF(F' + r + '=0,"",N' + r + '/F' + r + ')'
      ]);
    });
    sh.getRange(start, 1, values.length, 5).setNumberFormat('@');
    sh.getRange(start, 1, values.length, 10).setValues(values);
    sh.getRange(start, 11, formulas.length, 5).setFormulas(formulas);
    sh.getRange(start, 16, values.length, 1).setValues(values.map(function () { return [now]; }));
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function remove_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return { ok: true, deleted: deleteById_(sheet_(), String(id || '')) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Đọc ảnh bằng Gemini ---------------- */

const PROMPT = [
  'Đây là (các) ảnh chụp màn hình CHI TIẾT MỘT ĐƠN HÀNG trên Shopee Việt Nam. Nếu có nhiều ảnh thì chúng là các phần nối tiếp của cùng một đơn; đừng lặp lại sản phẩm xuất hiện ở cả hai ảnh.',
  'Hãy đọc và trả về DUY NHẤT một đối tượng JSON theo đúng cấu trúc sau:',
  '{"ma_don":"","ngay_dat":"","shop":"","shopee_voucher":0,"voucher_shop":0,"shopee_xu":0,"phi_van_chuyen_thuc":0,"thanh_tien":0,"items":[{"ten":"","phan_loai":"","sl":1,"gia_mua":0}]}',
  'Quy tắc:',
  '- ma_don: mã đơn hàng Shopee nếu thấy, không thấy thì để "".',
  '- ngay_dat: ngày đặt hàng dạng dd/mm/yyyy nếu thấy, không thấy thì để "".',
  '- shop: tên shop bán.',
  '- items: mỗi sản phẩm một phần tử. ten = tên sản phẩm; phan_loai = dòng "Phân loại hàng" (bỏ chữ "Phân loại hàng:"); sl = số lượng (số sau dấu x); gia_mua = giá thực mua của MỘT món, tức con số nằm cạnh giá bị gạch bỏ, KHÔNG lấy giá gạch bỏ.',
  '- shopee_voucher: số tiền giảm của dòng "Voucher từ Shopee" / "Shopee Voucher" (không tính giảm phí vận chuyển).',
  '- voucher_shop: số tiền giảm của dòng "Voucher từ Shop" / "Giảm giá của Shop".',
  '- shopee_xu: số tiền giảm do dùng Shopee Xu.',
  '- phi_van_chuyen_thuc: phí vận chuyển trừ đi giảm giá phí vận chuyển (không âm).',
  '- thanh_tien: dòng "Thành tiền" / tổng thanh toán của đơn.',
  '- Mọi số tiền là số nguyên đồng, dương, không có dấu chấm, dấu phẩy hay ký hiệu đ. Không thấy thì ghi 0.',
  '- Chỉ ghi những gì nhìn thấy trong ảnh, tuyệt đối không suy đoán hay bịa số.',
  '- KHÔNG trích tên người nhận, số điện thoại, địa chỉ hay bất kỳ thông tin cá nhân nào.'
].join('\n');

function extract_(images, props) {
  const key = String(props.getProperty('GEMINI_API_KEY') || '').trim();   // bỏ khoảng trắng/xuống dòng lỡ dán thừa
  if (!key) throw new Error('Chưa đặt GEMINI_API_KEY trong Script Properties.');
  if (!images.length) throw new Error('Chưa có ảnh.');
  const parts = [{ text: PROMPT }];
  images.forEach(function (im) {
    parts.push({ inline_data: { mime_type: im.mime || 'image/jpeg', data: im.data } });
  });
  const body = { contents: [{ parts: parts }], generationConfig: { responseMimeType: 'application/json' } };

  const r = geminiSmart_(key, body, String(props.getProperty('GEMINI_MODEL') || '').trim());
  const code = r.res.getResponseCode();
  const text = r.res.getContentText();
  if (code !== 200) throw new Error(explainGemini_(code, text) + (r.tried.length > 1 ? ' Đã thử: ' + r.tried.join(', ') + '.' : ''));

  const data = JSON.parse(text);
  const cand = data.candidates && data.candidates[0];
  if (!cand || !cand.content || !cand.content.parts) {
    const why = (cand && cand.finishReason) || (data.promptFeedback && data.promptFeedback.blockReason) || '';
    throw new Error('Gemini không trả về nội dung' + (why ? ' (' + why + ')' : '') + '. Thử lại hoặc dùng ảnh rõ hơn.');
  }
  let raw = cand.content.parts.filter(function (p) { return p.text && !p.thought; })
    .map(function (p) { return p.text; }).join('');
  raw = raw.replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim();
  const m = raw.match(/[\[{][\s\S]*[\]}]/);
  let o;
  try { o = JSON.parse(m ? m[0] : raw); }
  catch (e) { throw new Error('Không hiểu kết quả Gemini trả về, hãy thử lại: ' + raw.slice(0, 150)); }
  if (Array.isArray(o)) o = o[0] || {};

  return {
    maDon: String(o.ma_don || '').trim(),
    ngay: String(o.ngay_dat || '').trim(),
    shop: String(o.shop || '').trim(),
    sv: int_(o.shopee_voucher), vs: int_(o.voucher_shop), xu: int_(o.shopee_xu),
    ship: int_(o.phi_van_chuyen_thuc), thanhTien: int_(o.thanh_tien),
    items: (o.items || []).map(function (it) {
      return { ten: String(it.ten || '').trim(), phanLoai: String(it.phan_loai || '').trim(),
        sl: int_(it.sl) || 1, gia: int_(it.gia_mua) };
    })
  };
}

const DEFAULT_MODEL = 'gemini-flash-latest';
// Model dự phòng khi model chính quá tải (503) hoặc hết lượt (429). Lite đọc chữ trong ảnh vẫn tốt.
const FALLBACK_MODELS = ['gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-2.0-flash'];

/** Gọi Gemini, tự thử lại khi quá tải và tự đổi sang model dự phòng. Trả về {res, model, tried}. */
function geminiSmart_(key, body, preferred) {
  const list = [preferred || DEFAULT_MODEL].concat(FALLBACK_MODELS)
    .filter(function (m, i, a) { return m && a.indexOf(m) === i; });
  const tried = [];
  let res = null, model = list[0];
  for (let i = 0; i < list.length; i++) {
    model = list[i];
    for (let attempt = 0; attempt < 2; attempt++) {
      res = callGemini_(model, key, body);
      const c = res.getResponseCode();
      if (c === 200) return { res: res, model: model, tried: tried };
      if (attempt === 0 && (c === 500 || c === 503 || c === 504)) { Utilities.sleep(2500); continue; }  // bận tạm thời: đợi rồi thử lại 1 lần
      break;
    }
    const code = res.getResponseCode();
    tried.push(model + ' (' + code + ')');
    // Chỉ đổi model khi lỗi do model (quá tải, hết lượt, không tồn tại). Lỗi key/quyền thì đổi model cũng vô ích.
    if ([404, 429, 500, 503, 504].indexOf(code) === -1) break;
  }
  if (res.getResponseCode() === 404) {            // mọi tên model đều đã đổi: tự tìm model flash đang có
    model = pickModel_(key);
    res = callGemini_(model, key, body);
    if (res.getResponseCode() === 200) return { res: res, model: model, tried: tried };
    tried.push(model + ' (' + res.getResponseCode() + ')');
  }
  return { res: res, model: model, tried: tried };
}

function callGemini_(model, key, body) {
  return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
}

function pickModel_(key) {
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('Không lấy được danh sách model Gemini: ' + res.getContentText().slice(0, 200));
  const names = (JSON.parse(res.getContentText()).models || [])
    .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') !== -1; })
    .map(function (m) { return m.name.replace('models/', ''); })
    .filter(function (n) { return /flash/.test(n) && !/lite|tts|image|live|audio|embed|preview|exp/.test(n); })
    .sort().reverse();
  if (!names.length) throw new Error('Không tìm thấy model Gemini Flash nào. Hãy đặt GEMINI_MODEL trong Script Properties.');
  return names[0];
}

/** Dịch mã lỗi Gemini thành hướng dẫn dễ hiểu. */
function explainGemini_(code, text) {
  let msg = '';
  try { msg = (JSON.parse(text).error || {}).message || ''; } catch (e) { msg = String(text || '').slice(0, 200); }
  if (/API key not valid|API_KEY_INVALID/i.test(msg)) return 'API key Gemini sai hoặc đã bị xoá. Tạo key mới tại aistudio.google.com/app/apikey rồi dán lại vào GEMINI_API_KEY.';
  if (/location is not supported/i.test(msg)) return 'Gemini chưa hỗ trợ miễn phí ở khu vực của tài khoản/máy chủ này (' + msg + ').';
  if (code === 403) return 'Key bị từ chối (403): có thể key bị giới hạn, project chưa bật Generative Language API hoặc bị khoá. Thử tạo key mới trong AI Studio. Chi tiết: ' + msg;
  if (code === 429) return /limit: 0/.test(msg)
    ? 'Model này không có lượt miễn phí cho key của bạn (limit 0). Đặt GEMINI_MODEL = gemini-2.5-flash trong Script Properties, hoặc bật thanh toán cho project.'
    : 'Gemini báo hết lượt miễn phí tạm thời, thử lại sau ít phút.';
  if (code === 500 || code === 503 || code === 504) return 'Máy chủ Gemini của Google đang quá tải (' + code + '). Script đã tự thử lại và đổi model nhưng vẫn bận. Đợi 1–2 phút rồi bấm đọc lại.';
  if (code === 404) return 'Không tìm thấy model Gemini. Xoá GEMINI_MODEL trong Script Properties để script tự chọn model. Chi tiết: ' + msg;
  return 'Gemini lỗi ' + code + ': ' + msg;
}

/** Tự chẩn đoán: app gọi lệnh "diag" để biết hỏng ở bước nào. */
function diag_(props) {
  const checks = [];
  function add(ok, name, detail) { checks.push({ ok: ok, name: name, detail: detail || '' }); }
  function done() { return { ok: true, version: SCRIPT_VERSION, checks: checks }; }

  add(true, 'Apps Script', 'Đang chạy bản ' + SCRIPT_VERSION + ', mã PIN đúng.');
  try { sheet_(); add(true, 'Google Sheet', 'Ghi được vào sheet "' + SHEET_NAME + '".'); }
  catch (e) { add(false, 'Google Sheet', String(e && e.message || e)); }

  const raw = props.getProperty('GEMINI_API_KEY');
  const key = String(raw || '').trim();
  if (!key) { add(false, 'GEMINI_API_KEY', 'Chưa đặt trong Project Settings > Script Properties.'); return done(); }
  add(/^AIza/.test(key), 'GEMINI_API_KEY',
    'Đã đặt: ' + key.slice(0, 4) + '…' + key.slice(-4) + ' (' + key.length + ' ký tự).' +
    (raw !== key ? ' Có khoảng trắng thừa, script đã tự bỏ.' : '') +
    (/^AIza/.test(key) ? '' : ' Key Gemini thường bắt đầu bằng "AIza" và dài 39 ký tự: có thể bạn dán nhầm.'));

  const preferred = String(props.getProperty('GEMINI_MODEL') || '').trim();
  const body = { contents: [{ parts: [{ text: 'Trả về đúng JSON {"ok":true}' }] }],
    generationConfig: { responseMimeType: 'application/json' } };
  const t0 = Date.now();
  let r;
  try { r = geminiSmart_(key, body, preferred); }
  catch (e) {
    add(false, 'Quyền gọi Gemini', String(e && e.message || e) +
      ' → Trong trình soạn Apps Script, chọn hàm caiDat, bấm Chạy và cấp quyền, rồi Deploy > Manage deployments > Edit > New version.');
    return done();
  }
  const code = r.res.getResponseCode();
  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const via = r.tried.length ? ' (model trước bận/lỗi: ' + r.tried.join(', ') + ', đã tự chuyển)' : '';
  if (code === 200) add(true, 'Gọi thử Gemini', 'Model ' + r.model + ' trả lời sau ' + secs + ' giây' + via + '. Đọc ảnh sẽ hoạt động.');
  else add(false, 'Gọi thử Gemini', explainGemini_(code, r.res.getContentText()) + (r.tried.length ? ' Đã thử: ' + r.tried.join(', ') + '.' : ''));
  return done();
}

/** Chạy hàm này 1 lần trong trình soạn Apps Script để cấp quyền và tạo sheet. */
function caiDat() {
  sheet_();
  UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=1', { muteHttpExceptions: true });
}
