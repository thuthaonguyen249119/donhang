# donhang
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

const SHEET_NAME = 'Đơn hàng';
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
      case 'ping': return out_({ ok: true });
      case 'list': return out_({ ok: true, rows: listRows_() });
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
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Chưa đặt GEMINI_API_KEY trong Script Properties.');
  if (!images.length) throw new Error('Chưa có ảnh.');
  const parts = [{ text: PROMPT }];
  images.forEach(function (im) {
    parts.push({ inline_data: { mime_type: im.mime || 'image/jpeg', data: im.data } });
  });
  const body = { contents: [{ parts: parts }], generationConfig: { responseMimeType: 'application/json' } };

  let model = props.getProperty('GEMINI_MODEL') || 'gemini-flash-latest';
  let res = callGemini_(model, key, body);
  if (res.getResponseCode() === 404) {            // tên model đã đổi: tự tìm model flash đang có
    model = pickModel_(key);
    res = callGemini_(model, key, body);
  }
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code === 429) throw new Error('Gemini báo hết lượt miễn phí tạm thời, thử lại sau ít phút.');
  if (code !== 200) throw new Error('Gemini lỗi ' + code + ': ' + text.slice(0, 300));

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

/** Chạy hàm này 1 lần trong trình soạn Apps Script để cấp quyền và tạo sheet. */
function caiDat() {
  sheet_();
  UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=1', { muteHttpExceptions: true });
}
