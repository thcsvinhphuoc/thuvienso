/**
 * THƯ VIỆN SỐ – TRƯỜNG THCS VĨNH PHƯỚC
 * Máy chủ dữ liệu bằng Google Apps Script + Google Sheets
 *
 * CÁCH CÀI ĐẶT (làm 1 lần):
 * 1. Tạo 1 file Google Sheets mới (vd: "Du lieu Thu vien").
 * 2. Menu Tiện ích mở rộng > Apps Script. Xóa hết mã có sẵn, dán toàn bộ file này vào, bấm Lưu.
 * 3. Chọn hàm "setup" ở thanh công cụ > bấm Chạy > cấp quyền cho tài khoản Google.
 *    (Các trang tính DanhMuc, Sach, BanSach, BanDoc, MuonTra, CaiDat sẽ được tạo tự động.)
 * 4. Bấm Triển khai > Tùy chọn triển khai mới > loại "Ứng dụng web":
 *      - Thực thi dưới dạng: Tôi
 *      - Người có quyền truy cập: Bất kỳ ai
 *    Bấm Triển khai, sao chép URL (kết thúc bằng /exec).
 * 5. Mở index.html, dán URL vào dòng  const API_URL = '';
 *
 * Mật khẩu quản trị mặc định: admin123  (đổi trong trang Thiết lập).
 * Khi sửa mã này: Triển khai > Quản lý triển khai > Sửa (bút chì) > Phiên bản mới > Triển khai.
 */

const TABLES = {
  DanhMuc: ['id', 'ten', 'thuTu'],
  Sach: ['id', 'nhanDe', 'tacGia', 'nxb', 'namXB', 'gia', 'danhMucId', 'soPhanLoai', 'isbn', 'hashtag', 'anhBia', 'moTa', 'noiBat', 'ngayTao'],
  BanSach: ['dkcb', 'sachId', 'tinhTrang', 'ngayNhap', 'nguon', 'ghiChu'],
  BanDoc: ['ma', 'hoTen', 'loai', 'lop', 'gioiTinh', 'ngaySinh', 'sdt', 'ghiChu', 'pin'],
  MuonTra: ['id', 'dkcb', 'sachId', 'maBanDoc', 'ngayMuon', 'hanTra', 'ngayTra', 'trangThai', 'ghiChu', 'boiThuong'],
  CaiDat: ['key', 'value'],
  BaiViet: ['id', 'loai', 'tieuDe', 'tomTat', 'noiDung', 'anh', 'ngay', 'noiBat', 'hien'],
  TaiNguyen: ['id', 'loai', 'tieuDe', 'moTa', 'link', 'anh', 'mon', 'khoi', 'chuDe', 'tacGia', 'noiBat', 'hien', 'luotXem', 'ngayTao'],
  HoiDap: ['id', 'cauHoi', 'traLoi', 'thuTu'],
  LienKet: ['id', 'ten', 'link', 'anh', 'thuTu'],
  DanhGia: ['id', 'data', 'ngay'],
  MinhChung: ['id', 'nhom', 'ten', 'link', 'ngay', 'ghiChu'],
  DatMuon: ['id', 'maBanDoc', 'sachId', 'ngay', 'trangThai', 'ghiChu'],
  KiemKe: ['id', 'ten', 'ngayBatDau', 'ngayChot', 'trangThai', 'ghiChu', 'ketQua'],
  KiemKeCT: ['id', 'kiemKeId', 'dkcb', 'ngay'],
  ThanhLy: ['id', 'soBB', 'ngay', 'lyDo', 'ds', 'tongGia', 'ghiChu'],
  TruyCap: ['id', 'luot'],
  NhatKy: ['id', 'thoiGian', 'thietBi', 'noiDung'],
  TietDoc: ['id', 'ngay', 'lop', 'loai', 'mon', 'gv', 'soHS', 'noiDung']
};
const KEYS = { DanhMuc: 'id', Sach: 'id', BanSach: 'dkcb', BanDoc: 'ma', MuonTra: 'id', CaiDat: 'key',
  BaiViet: 'id', TaiNguyen: 'id', HoiDap: 'id', LienKet: 'id', DanhGia: 'id', MinhChung: 'id',
  DatMuon: 'id', KiemKe: 'id', KiemKeCT: 'id', ThanhLy: 'id', TruyCap: 'id', NhatKy: 'id', TietDoc: 'id' };
const DATA_VERSION = '5';

const DEFAULT_CFG = {
  matKhau: 'admin123',
  tenTruong: 'Trường THCS Vĩnh Phước',
  tienToDKCB: 'VP', soChuSoDKCB: '5',
  ngayMuonHS: '7', ngayMuonGV: '30', toiDaHS: '3', toiDaGV: '10'
};
const DEFAULT_DM = ['SÁCH GIÁO KHOA DÙNG CHUNG', 'SÁCH GIÁO KHOA', 'SÁCH GIÁO VIÊN', 'SÁCH THAM KHẢO', 'SÁCH THIẾU NHI', 'SÁCH TIẾNG ANH'];

/* ---------- Cài đặt ban đầu ---------- */
function setup() {
  PropertiesService.getScriptProperties().deleteProperty('ready');
  ensure_();
  SpreadsheetApp.getActive().toast('Đã tạo xong các trang tính dữ liệu thư viện.');
}

function ensure_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ready') === DATA_VERSION) return;
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABLES).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    const cols = TABLES[name];
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#d9f2e6');
      sh.setFrozenRows(1);
      if (name === 'CaiDat') {
        const rows = Object.keys(DEFAULT_CFG).map(function (k) { return [k, DEFAULT_CFG[k]]; });
        sh.getRange(2, 1, rows.length, 2).setValues(rows);
      }
      if (name === 'DanhMuc') {
        const rows = DEFAULT_DM.map(function (t, i) { return ['DM' + (i + 1), t, String(i + 1)]; });
        sh.getRange(2, 1, rows.length, 3).setValues(rows);
      }
    } else {
      // bổ sung cột mới nếu thiếu
      const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      cols.forEach(function (c) {
        if (head.indexOf(c) < 0) { sh.getRange(1, head.length + 1).setValue(c).setFontWeight('bold'); head.push(c); }
      });
    }
  });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch (e) {} }
  props.setProperty('ready', DATA_VERSION);
}

/* ---------- Đọc / ghi bảng ---------- */
function read_(name, withRow) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0].map(String);
  const key = KEYS[name];
  const tz = Session.getScriptTimeZone();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = {};
    head.forEach(function (h, j) {
      if (!h) return;
      let v = vals[i][j];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      o[h] = v === null || v === undefined ? '' : String(v);
    });
    if (withRow) o.__r = i + 1;
    if (o[key] !== '') out.push(o);
  }
  return out;
}

function write_(name, rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const data = rows.map(function (r) {
    return head.map(function (h) { return r[h] === undefined || r[h] === null ? '' : String(r[h]); });
  });
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, head.length).clearContent();
  if (data.length) {
    const rg = sh.getRange(2, 1, data.length, head.length);
    rg.setNumberFormat('@');
    rg.setValues(data);
  }
}

function cfg_() {
  const o = {};
  Object.keys(DEFAULT_CFG).forEach(function (k) { o[k] = DEFAULT_CFG[k]; });
  read_('CaiDat').forEach(function (r) { o[r.key] = r.value; });
  return o;
}

/* ---------- Xử lý thao tác ---------- */
function applyOps_(db, ops, meta) {
  meta = meta || {};
  ops.forEach(function (op) {
    if (op.t !== 'assert') return;
    const k = KEYS[op.table];
    const row = db[op.table].filter(function (r) { return String(r[k]) === String(op.key); })[0];
    if (!row || String(row[op.field]) !== String(op.equals))
      throw new Error(op.msg || 'Dữ liệu đã thay đổi, hãy bấm Tải lại.');
  });
  ops.forEach(function (op) {
    const k = KEYS[op.table];
    const m = meta[op.table] || (meta[op.table] = { changed: [], removed: false });
    if (op.t === 'upsert') {
      const idx = {};
      db[op.table].forEach(function (r, i) { idx[String(r[k])] = i; });
      (op.rows || []).forEach(function (row) {
        if (!row || row[k] === undefined || row[k] === '') throw new Error('Thiếu khóa ' + k + ' ở bảng ' + op.table);
        if (op.table === 'CaiDat' && row.key === 'matKhau' && !String(row.value || '').trim())
          throw new Error('Mật khẩu không được để trống');
        const clean = Object.assign({}, row); delete clean.__r;
        const i = idx[String(row[k])];
        if (i !== undefined) {
          const old = db[op.table][i];
          const nr = Object.assign({}, old, clean); nr.__r = old.__r;
          db[op.table][i] = nr;
          if (old.__r) m.changed.push(old.__r);
        } else { db[op.table].push(clean); idx[String(row[k])] = db[op.table].length - 1; }
      });
    } else if (op.t === 'remove') {
      const set = {};
      (op.keys || []).forEach(function (x) { set[String(x)] = 1; });
      if (op.table === 'CaiDat') delete set.matKhau;
      const before = db[op.table].length;
      db[op.table] = db[op.table].filter(function (r) { return !set[String(r[k])]; });
      if (db[op.table].length !== before) m.removed = true;
    }
  });
  return meta;
}

/* Ghi bảng tiết kiệm: chỉ ghi các dòng thay đổi và dòng mới; có xóa thì ghi lại cả bảng */
function flush_(name, rows, m) {
  if (m.removed) return write_(name, rows);
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const toRow = function (r) { return head.map(function (h) { return r[h] === undefined || r[h] === null ? '' : String(r[h]); }); };
  const byRow = {};
  rows.forEach(function (r) { if (r.__r) byRow[r.__r] = r; });
  const nums = m.changed.filter(function (n, i, a) { return a.indexOf(n) === i; }).sort(function (a, b) { return a - b; });
  let blocks = 0;
  for (let i = 0; i < nums.length; ) { let j = i; while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++; blocks++; i = j + 1; }
  if (blocks > 40) return write_(name, rows);
  for (let i = 0; i < nums.length; ) {
    let j = i; while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
    const data = [];
    for (let n = nums[i]; n <= nums[j]; n++) data.push(toRow(byRow[n]));
    const rg = sh.getRange(nums[i], 1, data.length, head.length);
    rg.setNumberFormat('@'); rg.setValues(data);
    i = j + 1;
  }
  const add = rows.filter(function (r) { return !r.__r; }).map(toRow);
  if (add.length) {
    const rg = sh.getRange(sh.getLastRow() + 1, 1, add.length, head.length);
    rg.setNumberFormat('@'); rg.setValues(add);
  }
}

/* Bộ nhớ đệm trang công khai (2 phút) để chịu được nhiều học sinh truy cập cùng lúc */
function cacheGet_() {
  try {
    const c = CacheService.getScriptCache(); const n = c.get('pub_n'); if (!n) return null;
    const keys = []; for (let i = 0; i < +n; i++) keys.push('pub_' + i);
    const m = c.getAll(keys); let s = '';
    for (let i = 0; i < keys.length; i++) { if (m[keys[i]] == null) return null; s += m[keys[i]]; }
    return s;
  } catch (e) { return null; }
}
function cachePut_(s) {
  try {
    const parts = {}; let n = 0;
    for (let i = 0; i < s.length; i += 30000) parts['pub_' + (n++)] = s.slice(i, i + 30000);
    if (n > 200) return;
    parts.pub_n = String(n);
    CacheService.getScriptCache().putAll(parts, 120);
  } catch (e) {}
}
function cacheClear_() { try { CacheService.getScriptCache().remove('pub_n'); } catch (e) {} }

function publicView_() {
  const cfg = cfg_();
  delete cfg.matKhau;
  const ban = read_('BanSach');
  const dem = {};
  let tongBan = 0;
  ban.forEach(function (b) {
    if (b.tinhTrang === 'Thanh lý' || b.tinhTrang === 'Mất') return;
    const c = dem[b.sachId] || (dem[b.sachId] = { tong: 0, san: 0 });
    c.tong++; tongBan++;
    if (b.tinhTrang === 'Sẵn sàng') c.san++;
  });
  const sh = SpreadsheetApp.getActive();
  return {
    ok: true,
    DanhMuc: read_('DanhMuc'),
    Sach: read_('Sach'),
    BaiViet: read_('BaiViet').filter(function (r) { return r.hien !== '0'; }),
    TaiNguyen: read_('TaiNguyen').filter(function (r) { return r.hien !== '0'; }),
    HoiDap: read_('HoiDap'),
    LienKet: read_('LienKet'),
    dem: dem,
    cfg: cfg,
    stats: {
      banSach: tongBan,
      banDoc: Math.max(0, sh.getSheetByName('BanDoc').getLastRow() - 1),
      luotMuon: Math.max(0, sh.getSheetByName('MuonTra').getLastRow() - 1)
    }
  };
}

/* Tăng lượt xem tài nguyên số (gọi công khai, chỉ cộng 1) */
function viewInc_(id) {
  if (!id) return { ok: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('TaiNguyen');
    const vals = sh.getDataRange().getValues();
    const head = vals[0].map(String);
    const ci = head.indexOf('id'), cv = head.indexOf('luotXem');
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][ci]) === id) {
        const n = (parseInt(vals[i][cv], 10) || 0) + 1;
        sh.getRange(i + 1, cv + 1).setNumberFormat('@').setValue(String(n));
        return { ok: true, luotXem: n };
      }
    }
    return { ok: false };
  } finally { lock.releaseLock(); }
}

function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

/* Đếm lượt truy cập website theo ngày */
function hit_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = today_();
    const sh = SpreadsheetApp.getActive().getSheetByName('TruyCap');
    const vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === d) {
        sh.getRange(i + 1, 2).setNumberFormat('@').setValue(String((parseInt(vals[i][1], 10) || 0) + 1));
        return { ok: true };
      }
    }
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 2).setNumberFormat('@').setValues([[d, '1']]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* Cổng bạn đọc: đăng nhập bằng mã bạn đọc + mật khẩu (mặc định: ngày sinh ddmmyyyy, nếu không có thì là mã bạn đọc) */
function pinOk_(r, pin) {
  pin = String(pin || '').trim();
  if (!pin) return false;
  if (r.pin) return pin === String(r.pin);
  const m = String(r.ngaySinh || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return pin === m[3] + m[2] + m[1];
  return pin.toUpperCase() === String(r.ma).toUpperCase();
}
function reader_(req) {
  const ma = String(req.ma || '').trim().toUpperCase();
  const bd = read_('BanDoc');
  const r = bd.filter(function (b) { return String(b.ma).toUpperCase() === ma; })[0];
  if (!r || !pinOk_(r, req.pin)) { Utilities.sleep(1500); return { ok: false, error: 'Sai mã bạn đọc hoặc mật khẩu' }; }
  const cfg = cfg_();
  const sub = req.sub || 'info';
  if (sub !== 'info') {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (sub === 'datmuon') {
        const dm = read_('DatMuon');
        const open = dm.filter(function (d) { return d.maBanDoc === r.ma && d.trangThai === 'Chờ'; });
        const max = parseInt(r.loai === 'GV' ? cfg.toiDaGV : cfg.toiDaHS, 10) || 3;
        if (open.length >= max) return { ok: false, error: 'Bạn chỉ được đặt trước tối đa ' + max + ' cuốn cùng lúc' };
        if (open.some(function (d) { return d.sachId === req.sachId; })) return { ok: false, error: 'Bạn đã đặt mượn sách này rồi' };
        if (!read_('Sach').some(function (s) { return s.id === req.sachId; })) return { ok: false, error: 'Không tìm thấy sách' };
        dm.push({ id: 'D' + new Date().getTime().toString(36).toUpperCase() + Math.floor(Math.random() * 1000), maBanDoc: r.ma, sachId: String(req.sachId),
          ngay: today_(), trangThai: 'Chờ', ghiChu: String(req.ghiChu || '').slice(0, 200) });
        write_('DatMuon', dm);
      } else if (sub === 'huy') {
        const dm = read_('DatMuon');
        const d = dm.filter(function (x) { return x.id === req.id && x.maBanDoc === r.ma && x.trangThai === 'Chờ'; })[0];
        if (!d) return { ok: false, error: 'Không hủy được yêu cầu này' };
        d.trangThai = 'Đã hủy';
        write_('DatMuon', dm);
      } else if (sub === 'doipin') {
        const p = String(req.pinMoi || '').trim();
        if (p.length < 4) return { ok: false, error: 'Mật khẩu mới tối thiểu 4 ký tự' };
        r.pin = p;
        write_('BanDoc', bd);
      }
    } finally { lock.releaseLock(); }
  }
  const pick = function (o, ks) { const x = {}; ks.forEach(function (k) { x[k] = o[k]; }); return x; };
  return {
    ok: true,
    ban: pick(r, ['ma', 'hoTen', 'loai', 'lop']),
    muon: read_('MuonTra').filter(function (m) { return m.maBanDoc === r.ma; })
      .map(function (m) { return pick(m, ['id', 'dkcb', 'sachId', 'ngayMuon', 'hanTra', 'ngayTra', 'trangThai']); }),
    dat: read_('DatMuon').filter(function (d) { return d.maBanDoc === r.ma; })
  };
}

function allData_() {
  const o = {};
  Object.keys(TABLES).forEach(function (t) { if (t !== 'NhatKy') o[t] = read_(t); });
  o.CaiDat = o.CaiDat.filter(function (r) { return r.key !== 'matKhau'; });
  return o;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Điểm vào Web App ---------- */
function doGet(e) {
  try {
    ensure_();
    const a = (e && e.parameter && e.parameter.action) || 'public';
    if (a === 'public') {
      let str = cacheGet_();
      if (!str) { str = JSON.stringify(publicView_()); cachePut_(str); }
      return ContentService.createTextOutput(str).setMimeType(ContentService.MimeType.JSON);
    }
    if (a === 'view') return json_(viewInc_(String(e.parameter.id || '')));
    if (a === 'hit') return json_(hit_());
    return json_({ ok: true, msg: 'Thư viện số API đang hoạt động' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    ensure_();
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'reader') return json_(reader_(req));
    const cfg = cfg_();
    if (String(req.pass || '') !== String(cfg.matKhau)) {
      Utilities.sleep(1500); // làm chậm việc dò mật khẩu
      return json_({ ok: false, auth: false, error: 'Sai mật khẩu quản trị' });
    }
    if (req.action === 'sheet') {
      // Đọc dữ liệu từ một file Google Sheets khác (để nhập sách, bạn đọc, phiếu mượn…)
      const url = String(req.url || '');
      if (!/docs\.google\.com\/spreadsheets\/d\//.test(url)) return json_({ ok: false, error: 'Đường liên kết Google Sheets không hợp lệ' });
      let ss;
      try { ss = SpreadsheetApp.openByUrl(url); }
      catch (err) { return json_({ ok: false, error: 'Không mở được file: tài khoản Google chủ của hệ thống thư viện chưa có quyền xem file này. Hãy chia sẻ file cho tài khoản đó, hoặc dùng cách "Dán dữ liệu".' }); }
      const m = url.match(/[#&?]gid=(\d+)/);
      let sh = ss.getSheets()[0];
      if (m) ss.getSheets().forEach(function (x) { if (String(x.getSheetId()) === m[1]) sh = x; });
      return json_({ ok: true, name: sh.getName(), values: sh.getDataRange().getDisplayValues() });
    }
    if (req.action === 'log') {
      const sh = SpreadsheetApp.getActive().getSheetByName('NhatKy');
      const last = sh.getLastRow(), n = Math.min(500, Math.max(0, last - 1));
      const vals = n ? sh.getRange(last - n + 1, 1, n, 4).getValues() : [];
      return json_({ ok: true, data: vals.map(function (v) { return { id: String(v[0]), thoiGian: String(v[1]), thietBi: String(v[2]), noiDung: String(v[3]) }; }) });
    }

    if (req.action === 'login') return json_({ ok: true });
    if (req.action === 'all') return json_({ ok: true, data: allData_() });

    if (req.action === 'batch') {
      const ops = req.ops || [];
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try {
        const db = {}, touched = {};
        ops.forEach(function (op) {
          if (!TABLES[op.table] || op.table === 'NhatKy') throw new Error('Bảng không hợp lệ: ' + op.table);
          if (!db[op.table]) db[op.table] = read_(op.table, true);
          if (op.t !== 'assert') touched[op.table] = 1;
        });
        const meta = applyOps_(db, ops, {});
        Object.keys(touched).forEach(function (t) { flush_(t, db[t], meta[t] || { changed: [], removed: false }); });
        try {
          const sum = ops.filter(function (o) { return o.t !== 'assert'; }).map(function (o) { return o.table + (o.t === 'remove' ? ' -' + (o.keys || []).length : ' ' + (o.rows || []).length); }).join(', ');
          const tz = Session.getScriptTimeZone();
          SpreadsheetApp.getActive().getSheetByName('NhatKy').appendRow([
            'L' + new Date().getTime().toString(36), Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
            String(req.who || '').slice(0, 60), (String(req.note || '') + ' [' + sum + ']').slice(0, 400)]);
        } catch (e) {}
        cacheClear_();
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'Thao tác không hợp lệ' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
