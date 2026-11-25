/**
 * TIC · Penghubung aplikasi (PWA) dengan Google Sheets
 * Tambahkan sebagai file baru "Api" di proyek Apps Script Sheets TIC (Ekstensi > Apps Script > + > Skrip).
 * Lalu: Terapkan > Penerapan baru > Jenis: Aplikasi web
 *   Jalankan sebagai: Saya | Yang memiliki akses: Siapa saja
 */
const KODE_API = 'TIC2026'; // ganti dengan kode admin Anda (boleh sama dengan KODE_ADMIN di skrip Form)

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function ymd_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v || ''); }

function baca_() {
  const ss = SpreadsheetApp.getActive(), P = PropertiesService.getScriptProperties();
  const mem = ss.getSheetByName('Anggota').getRange('A2:A201').getValues().flat()
    .map(s => String(s).trim()).filter(Boolean);
  const mt = ss.getSheetByName('Pertandingan').getRange('A2:G501').getValues()
    .filter(r => String(r[1]).trim() !== '')
    .map(r => [ymd_(r[0]), String(r[1]).trim(), String(r[2]).trim(), String(r[3]).trim(), String(r[4]).trim(), Number(r[5]), Number(r[6])]);
  const js = ss.getSheetByName('Jadwal');
  let jd = [];
  if (js && js.getLastRow() > 1) {
    jd = js.getRange(2, 1, js.getLastRow() - 1, 8).getValues()
      .filter(r => r[1] !== '' && r[1] !== null)
      .map(r => ({
        tgl: ymd_(r[1]),
        jam: String(r[2]).replace(/\s*WIB/i, '').replace('.', ':').trim(),
        tempat: String(r[3]), cat: String(r[4]),
        peserta: String(r[6]).split(',').map(s => s.trim()).filter(Boolean)
      }));
  }
  const raw = JSON.stringify([mem, mt, jd]);
  const v = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw));
  return { v: v, updated: P.getProperty('UPD') || null, admins: JSON.parse(P.getProperty('ADM') || '[]'),
           members: mem, matches: mt, jadwal: jd };
}

function valid_(d) {
  const mem = new Set(d.members || []);
  for (const m of (d.matches || [])) {
    const n = [m[1], m[2], m[3], m[4]];
    if (new Set(n).size < 4 || n.some(x => !mem.has(x))) return 'pertandingan tidak valid';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(m[0])) return 'tanggal tidak valid';
    if (!Number.isInteger(Number(m[5])) || !Number.isInteger(Number(m[6])) || m[5] < 0 || m[6] < 0 || m[5] > 9 || m[6] > 9) return 'skor tidak valid';
  }
  return '';
}

function simpan_(d) {
  const ss = SpreadsheetApp.getActive(), tz = tz_();
  const a = ss.getSheetByName('Anggota');
  a.getRange('A2:A201').clearContent();
  const mem = (d.members || []).slice(0, 200).map(n => [String(n).trim()]);
  if (mem.length) a.getRange(2, 1, mem.length, 1).setValues(mem);
  const p = ss.getSheetByName('Pertandingan');
  p.getRange('A2:G501').clearContent();
  const ms = (d.matches || []).slice(0, 500).map(m =>
    [Utilities.parseDate(m[0], tz, 'yyyy-MM-dd'), m[1], m[2], m[3], m[4], Number(m[5]), Number(m[6])]);
  if (ms.length) p.getRange(2, 1, ms.length, 7).setValues(ms);
  const j = ss.getSheetByName('Jadwal');
  if (j) {
    const js = d.jadwal || [], n = js.length;
    j.getRange('B2:E200').clearContent();
    j.getRange('G2:G200').clearContent();
    if (n) {
      j.getRange(2, 2, n, 4).setValues(js.map(x => [Utilities.parseDate(x.tgl, tz, 'yyyy-MM-dd'),
        String(x.jam).replace(':', '.') + ' WIB', x.tempat || '', x.cat || '']));
      j.getRange(2, 2, n, 1).setNumberFormat('dd-mmm-yyyy');
      j.getRange(2, 7, n, 1).setValues(js.map(x => [(x.peserta || []).join(', ')]));
      if (n > 1) ['A', 'F', 'H'].forEach(c =>
        j.getRange(c + '2').copyTo(j.getRange(c + '3:' + c + (n + 1)), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false));
    }
  }
  const P = PropertiesService.getScriptProperties();
  P.setProperty('UPD', new Date().toISOString());
  P.setProperty('ADM', JSON.stringify(d.admins || []));
}

// Hapus otomatis jadwal yang tanggalnya sudah lewat
function bersihkanJadwal_() {
  const js = SpreadsheetApp.getActive().getSheetByName('Jadwal');
  if (!js || js.getLastRow() < 2) return;
  const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  const rows = js.getRange(2, 1, js.getLastRow() - 1, 8).getValues().filter(r => r[1] !== '' && r[1] !== null);
  const keep = rows.filter(r => ymd_(r[1]) >= today);
  if (keep.length === rows.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    js.getRange('B2:E200').clearContent();
    js.getRange('G2:G200').clearContent();
    if (keep.length) {
      js.getRange(2, 2, keep.length, 4).setValues(keep.map(r => [r[1], r[2], r[3], r[4]]));
      js.getRange(2, 2, keep.length, 1).setNumberFormat('dd-mmm-yyyy');
      js.getRange(2, 7, keep.length, 1).setValues(keep.map(r => [r[6]]));
      if (keep.length > 1) ['A', 'F', 'H'].forEach(c =>
        js.getRange(c + '2').copyTo(js.getRange(c + '3:' + c + (keep.length + 1)), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false));
    }
  } finally { lock.releaseLock(); }
}

// Anggota mendaftar atau membatalkan diri sendiri pada sebuah jadwal (tanpa kode admin)
function ikut_(b) {
  const ss = SpreadsheetApp.getActive();
  const nama = String(b.nama || '').trim();
  const mem = ss.getSheetByName('Anggota').getRange('A2:A201').getValues().flat().map(s => String(s).trim());
  if (!nama || mem.indexOf(nama) < 0) return { error: 'nama tidak dikenal' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const js = ss.getSheetByName('Jadwal');
    if (!js || js.getLastRow() < 2) return { error: 'jadwal tidak ditemukan' };
    const rows = js.getRange(2, 1, js.getLastRow() - 1, 8).getValues();
    const idx = rows.findIndex(r => ymd_(r[1]) === b.tgl &&
      String(r[2]).replace(/\s*WIB/i, '').replace('.', ':').trim() === b.jam &&
      String(r[3]) === String(b.tempat || ''));
    if (idx < 0) return { error: 'jadwal tidak ditemukan' };
    let p = String(rows[idx][6]).split(',').map(s => s.trim()).filter(Boolean);
    const ada = p.indexOf(nama) >= 0;
    if (b.ikut && !ada) p.push(nama);
    if (!b.ikut && ada) p = p.filter(x => x !== nama);
    js.getRange(idx + 2, 7).setValue(p.join(', '));
    return baca_();
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  bersihkanJadwal_();
  return out_(baca_());
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ error: 'format' }); }
  bersihkanJadwal_();
  if (b.action === 'ikut') return out_(ikut_(b));
  if (b.code !== KODE_API) return out_({ error: 'kode' });
  if (b.action === 'cek') return out_({ ok: true });
  if (b.action !== 'simpan') return out_({ error: 'aksi' });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (b.v !== baca_().v) return out_({ error: 'versi' });
    const err = valid_(b.data || {});
    if (err) return out_({ error: err });
    simpan_(b.data);
    return out_(baca_());
  } finally { lock.releaseLock(); }
}
