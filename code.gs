/**
 * ============================================================
 *  ARISAN BERJAYA — code.gs (Backend)
 *  PT. Berjaya Makmur Indonesia
 *  Runtime: V8 | Zona waktu project: Asia/Jakarta
 * ============================================================
 */

// ==================== KONFIGURASI ====================
var TZ = 'Asia/Jakarta';
var SESSION_TTL = 21600;      // 6 jam (batas maksimal CacheService)
var MAX_GAGAL_LOGIN = 5;      // maksimal salah PIN sebelum dikunci
var DURASI_KUNCI = 900;       // 15 menit
var PAGE_SIZE = 20;           // default limit pagination
var KOLOM_ANGKA = ['Nominal_Setoran', 'Jumlah_Slot', 'Putaran_Aktif', 'No_Urut', 'Putaran_Menang', 'Putaran_Ke', 'Nominal', 'Total_Dana'];
var PRIORITAS_STATUS = { 'Lunas': 3, 'Menunggu Verifikasi': 2, 'Ditolak': 1 };
var KUNCI_PENGATURAN = ['Nama_Aplikasi', 'Nama_Perusahaan', 'Kontak_Admin', 'Nama_Bank', 'No_Rekening', 'Atas_Nama', 'Quote'];
var MAX_B64 = 7000000;        // ±5 MB file asli setelah base64
var FOLDER_ROOT = 'Arisan Berjaya';
var FOLDER = { foto: 'Foto Profil', struk: 'Struk Setoran', kwitansi: 'Kwitansi' };

// ==================== ENTRY POINT ====================
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Arisan Berjaya')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fondasi modular: <?!= include('namafile'); ?>
function include(namaFile) {
  return HtmlService.createHtmlOutputFromFile(namaFile).getContent();
}

// ==================== HELPER UMUM ====================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'); }
function fmtL_(d, pola) { return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), pola) : ''; }
function pad4_(n) { return ('0000' + n).slice(-4); }

// Konversi string angka (hasil getDisplayValues) ke integer
function num_(v) {
  var n = parseInt(String(v == null ? '' : v).replace(/[^\d-]/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

// Bersihkan input teks + cegah formula injection
function clean_(v, maks) {
  var s = String(v == null ? '' : v).trim().replace(/^[=+@]+/, '');
  return maks ? s.slice(0, maks) : s;
}

// Normalisasi No HP: hanya angka, 62xxx -> 0xxx
function normHp_(v) {
  var s = String(v || '').replace(/\D/g, '');
  if (s.indexOf('62') === 0) s = '0' + s.slice(2);
  return s;
}

// Parse 'dd/MM/yyyy' atau 'dd/MM/yyyy HH:mm:ss' menjadi Date
function parseTgl_(s) {
  var m = String(s || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
}

function hariIni_() { return parseTgl_(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy')); }
function tglSaja_(d) { return d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null; }
function selisihHari_(a, b) { return Math.round((a.getTime() - b.getTime()) / 86400000); }

// 'yyyy-MM-dd' (input date HTML) -> 'dd/MM/yyyy'
function ymdKeDmy_(s) {
  var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
}

function tambahPeriode_(d, periode, n) {
  var x = new Date(d.getTime());
  if (periode === 'Mingguan') x.setDate(x.getDate() + 7 * n);
  else x.setMonth(x.getMonth() + n);
  return x;
}

// Jatuh tempo (sekaligus jadwal undian) untuk putaran ke-n
function jatuhTempo_(g, putaran) {
  var d = parseTgl_(g.Tanggal_Mulai);
  return d ? tambahPeriode_(d, g.Periode, Math.max(0, putaran - 1)) : null;
}

// Wrapper standar: semua fungsi publik dibungkus try-catch
function run_(fn) {
  try {
    var d = fn();
    return { success: true, data: d === undefined ? null : d };
  } catch (e) {
    return { success: false, message: e && e.message ? e.message : String(e) };
  }
}

// Kunci proses tulis agar tidak bentrok (ID ganda, undian ganda)
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ==================== HASH PIN ====================
function hashPin_(pin, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(pin), Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}
function newSalt_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }

// ==================== DATA LAYER (BATCH) ====================
function getTable_(nama) {
  var sh = ss_().getSheetByName(nama);
  if (!sh) throw new Error('Sheet "' + nama + '" belum ada. Jalankan setupDatabase() di setup.gs terlebih dahulu.');
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) throw new Error('Header sheet "' + nama + '" kosong. Jalankan setupDatabase().');
  var values = sh.getRange(1, 1, Math.max(lastRow, 1), lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idx = {};
  headers.forEach(function (h, i) { if (h) idx[h] = i; });
  return { sheet: sh, headers: headers, idx: idx, rows: values.slice(1) };
}

function rowObj_(t, i) {
  var r = t.rows[i];
  var o = { _i: i };
  for (var j = 0; j < t.headers.length; j++) if (t.headers[j]) o[t.headers[j]] = r[j];
  return o;
}

// Semua baris -> array object (baris kosong dilewati)
function objs_(t) {
  var out = [];
  for (var i = 0; i < t.rows.length; i++) {
    if (String(t.rows[i][0]).trim() === '') continue;
    out.push(rowObj_(t, i));
  }
  return out;
}

function findIdx_(t, kolom, nilai) {
  var c = t.idx[kolom];
  if (c === undefined) return -1;
  for (var i = 0; i < t.rows.length; i++) if (String(t.rows[i][c]) === String(nilai)) return i;
  return -1;
}

// Tambah banyak baris sekaligus (1x setValues)
function appendObjs_(t, objs) {
  if (!objs || !objs.length) return;
  var sh = t.sheet;
  var rows = objs.map(function (o) {
    return t.headers.map(function (h) { return o[h] !== undefined && o[h] !== null ? o[h] : ''; });
  });
  var formats = rows.map(function () {
    return t.headers.map(function (h) { return KOLOM_ANGKA.indexOf(h) > -1 ? '0' : '@'; });
  });
  var start = sh.getLastRow() + 1;
  var butuh = start + rows.length - 1;
  if (butuh > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), butuh - sh.getMaxRows() + 100);
  var rg = sh.getRange(start, 1, rows.length, t.headers.length);
  rg.setNumberFormats(formats);
  rg.setValues(rows);
  rows.forEach(function (r) { t.rows.push(r.map(String)); });
}

// Update satu baris (1x setValues untuk 1 baris)
function updateObj_(t, i, perubahan) {
  var row = t.rows[i].slice();
  Object.keys(perubahan).forEach(function (k) {
    if (t.idx[k] !== undefined) row[t.idx[k]] = perubahan[k];
  });
  t.sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
  t.rows[i] = row.map(String);
  return rowObj_(t, i);
}

function kelompok_(arr, key) {
  var m = {};
  arr.forEach(function (o) { (m[o[key]] = m[o[key]] || []).push(o); });
  return m;
}

function paginate_(arr, p) {
  var off = Math.max(0, num_(p && p.offset));
  var lim = Math.min(100, num_(p && p.limit) || PAGE_SIZE);
  return { items: arr.slice(off, off + lim), total: arr.length, offset: off, hasMore: off + lim < arr.length };
}

function cocok_(q, fields) {
  if (!q) return true;
  return fields.join(' ').toLowerCase().indexOf(q) > -1;
}

// ID master sekuensial: USR-0001, GRP-0001, AGT-0001
function nextMasterId_(t, kolom, prefix) {
  var max = 0, c = t.idx[kolom], re = new RegExp('^' + prefix + '-(\\d+)$');
  t.rows.forEach(function (r) { var m = String(r[c]).match(re); if (m) max = Math.max(max, +m[1]); });
  return prefix + '-' + pad4_(max + 1);
}

// ID transaksi sekuensial reset harian: STR-ddMMyy-0001
function nextDailyId_(t, kolom, prefix) {
  var tag = Utilities.formatDate(new Date(), TZ, 'ddMMyy');
  var max = 0, c = t.idx[kolom], re = new RegExp('^' + prefix + '-' + tag + '-(\\d+)$');
  t.rows.forEach(function (r) { var m = String(r[c]).match(re); if (m) max = Math.max(max, +m[1]); });
  return prefix + '-' + tag + '-' + pad4_(max + 1);
}

// ==================== SESSION & AUTH ====================
function cache_() { return CacheService.getScriptCache(); }

function createSession_(u) {
  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  var data = { id: u.ID_User, nama: u.Nama, role: u.Role, hp: u.No_HP, foto: u.Foto_Profil || '' };
  cache_().put('S_' + token, JSON.stringify(data), SESSION_TTL);
  return { token: token, user: data };
}

// ==================== GOOGLE DRIVE (FOTO, STRUK, KWITANSI) ====================
// Ambil/buat folder, ID disimpan di Script Properties agar tidak mencari ulang
function folderDari_(kunci, nama, parent) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(kunci);
  if (id) {
    try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* folder hilang, buat ulang */ }
  }
  var it = parent ? parent.getFoldersByName(nama) : DriveApp.getFoldersByName(nama);
  var folder = it.hasNext() ? it.next() : (parent ? parent.createFolder(nama) : DriveApp.createFolder(nama));
  props.setProperty(kunci, folder.getId());
  return folder;
}
function rootFolder_() { return folderDari_('FOLDER_ROOT', FOLDER_ROOT, null); }
function folder_(nama) { return folderDari_('FOLDER_' + nama, nama, rootFolder_()); }
function folderIds_() {
  var m = {};
  Object.keys(FOLDER).forEach(function (k) { m[folder_(FOLDER[k]).getId()] = k; });
  return m;
}

// Bagikan "siapa saja yang punya link" (bisa gagal di akun Workspace yang membatasi)
function bagikan_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return true; } catch (e) { return false; }
}
function hapusFile_(id) {
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* abaikan */ }
}
function gantiNama_(id, nama) {
  if (!id) return;
  try { DriveApp.getFileById(id).setName(nama); } catch (e) { /* abaikan */ }
}

// Simpan gambar base64 (JPEG/PNG) ke folder Drive, return File ID
function simpanGambar_(dataUrl, nama, namaFolder) {
  var b64 = String(dataUrl || '').replace(/^data:[^,]*,/, '');
  if (!b64) throw new Error('File gambar kosong.');
  if (b64.length > MAX_B64) throw new Error('Ukuran gambar maksimal 5 MB.');
  var bytes;
  try { bytes = Utilities.base64Decode(b64); } catch (e) { throw new Error('Format gambar tidak valid.'); }
  // Validasi magic bytes: JPEG (FF D8) atau PNG (89 50)
  var jpg = bytes.length > 2 && bytes[0] === -1 && bytes[1] === -40;
  var png = bytes.length > 2 && bytes[0] === -119 && bytes[1] === 80;
  if (!jpg && !png) throw new Error('File harus berupa gambar JPG atau PNG.');
  var blob = Utilities.newBlob(bytes, jpg ? 'image/jpeg' : 'image/png', nama + (jpg ? '.jpg' : '.png'));
  var file = folder_(namaFolder).createFile(blob);
  bagikan_(file);
  return file.getId();
}

function updateSesi_(token, perubahan) {
  var raw = cache_().get('S_' + token);
  if (!raw) return;
  var s = JSON.parse(raw);
  Object.keys(perubahan).forEach(function (k) { s[k] = perubahan[k]; });
  cache_().put('S_' + token, JSON.stringify(s), SESSION_TTL);
}

// Map ID_User -> File ID foto profil
function fotoMap_(users) {
  var m = {};
  (users || objs_(getTable_('Users'))).forEach(function (u) { if (u.Foto_Profil) m[u.ID_User] = u.Foto_Profil; });
  return m;
}

function auth_(token, role) {
  if (!token) throw new Error('SESI_HABIS');
  var raw = cache_().get('S_' + token);
  if (!raw) throw new Error('SESI_HABIS');
  cache_().put('S_' + token, raw, SESSION_TTL); // perpanjang sesi
  var s = JSON.parse(raw);
  if (role && s.role !== role) throw new Error('Akses ditolak: fitur ini khusus ' + role + '.');
  return s;
}

// Versi cache dashboard (di-bump setiap ada perubahan data)
function bump_() { cache_().put('DASH_VER', String(new Date().getTime()), 21600); }
function dashVer_() { return cache_().get('DASH_VER') || '0'; }

function log_(sess, aksi, detail) {
  var t = getTable_('Log_Aktivitas');
  appendObjs_(t, [{ Timestamp: now_(), ID_User: sess.id, Nama: sess.nama, Role: sess.role, Aksi: aksi, Detail: clean_(detail, 300) }]);
}

function pubUser_(o) {
  var u = {};
  Object.keys(o).forEach(function (k) { if (k !== 'PIN_Hash' && k !== 'Salt' && k !== '_i') u[k] = o[k]; });
  return u;
}

function getSettings_() {
  var m = {};
  try {
    objs_(getTable_('Pengaturan')).forEach(function (r) { m[r.Key] = r.Value; });
  } catch (e) { /* sheet belum ada */ }
  return m;
}

function validPin_(pin) {
  if (!/^\d{6}$/.test(String(pin || ''))) throw new Error('PIN harus 6 digit angka.');
}

// ==================== API: LOGIN ====================
function login(identitas, pin) {
  return run_(function () {
    identitas = String(identitas || '').trim();
    pin = String(pin || '').trim();
    if (!identitas) throw new Error('No. HP / ID Peserta wajib diisi.');
    validPin_(pin);

    var kunci = 'GAGAL_' + (normHp_(identitas) || identitas.toUpperCase());
    var gagal = num_(cache_().get(kunci));
    if (gagal >= MAX_GAGAL_LOGIN) throw new Error('Akun dikunci 15 menit karena 5x salah PIN. Coba lagi nanti.');

    var t = getTable_('Users');
    var users = objs_(t);
    var hp = normHp_(identitas);
    var u = null;
    for (var i = 0; i < users.length; i++) {
      var x = users[i];
      if (String(x.ID_User).toUpperCase() === identitas.toUpperCase() || (hp.length >= 9 && normHp_(x.No_HP) === hp)) { u = x; break; }
    }

    if (!u || hashPin_(pin, u.Salt) !== u.PIN_Hash) {
      gagal++;
      cache_().put(kunci, String(gagal), DURASI_KUNCI);
      var sisa = MAX_GAGAL_LOGIN - gagal;
      throw new Error(sisa > 0 ? 'No. HP/ID atau PIN salah. Sisa percobaan: ' + sisa + '.' : 'Akun dikunci 15 menit karena 5x salah PIN.');
    }
    if (u.Status !== 'Aktif') throw new Error('Akun kamu nonaktif. Hubungi admin.');

    cache_().remove(kunci);
    updateObj_(t, u._i, { Terakhir_Login: now_() });
    var sesi = createSession_(u);
    log_(sesi.user, 'Login', u.Nama + ' masuk ke aplikasi');
    SpreadsheetApp.flush();
    return sesi;
  });
}

function checkSession(token) {
  return run_(function () { return auth_(token); });
}

function logout(token) {
  return run_(function () { if (token) cache_().remove('S_' + token); return true; });
}

function changePin(token, pinLama, pinBaru) {
  return run_(function () {
    var s = auth_(token);
    validPin_(pinBaru);
    return withLock_(function () {
      var t = getTable_('Users');
      var i = findIdx_(t, 'ID_User', s.id);
      if (i < 0) throw new Error('User tidak ditemukan.');
      var u = rowObj_(t, i);
      if (hashPin_(String(pinLama || ''), u.Salt) !== u.PIN_Hash) throw new Error('PIN lama salah.');
      var salt = newSalt_();
      updateObj_(t, i, { PIN_Hash: hashPin_(pinBaru, salt), Salt: salt });
      log_(s, 'Ganti PIN', s.nama + ' mengganti PIN');
      SpreadsheetApp.flush();
      return true;
    });
  });
}

// ==================== HELPER STATUS & GRUP ====================
function statusMap_(setoran) {
  var m = {};
  setoran.forEach(function (s) {
    var k = s.ID_Grup + '|' + s.ID_User + '|' + num_(s.Putaran_Ke);
    var cur = m[k];
    if (!cur || (PRIORITAS_STATUS[s.Status] || 0) > (PRIORITAS_STATUS[cur] || 0)) m[k] = s.Status;
  });
  return m;
}

function grupCard_(g, ang, smap, today) {
  var slot = num_(g.Jumlah_Slot) || 1;
  var p = Math.max(1, num_(g.Putaran_Aktif));
  var selesai = g.Status === 'Selesai';
  var pt = Math.min(p, slot);
  var due = jatuhTempo_(g, pt);
  var lunas = 0;
  ang.forEach(function (a) { if (smap[g.ID_Grup + '|' + a.ID_User + '|' + p] === 'Lunas') lunas++; });
  var done = selesai ? slot : Math.min(p - 1, slot);
  var nominal = num_(g.Nominal_Setoran);
  return {
    ID_Grup: g.ID_Grup, Nama_Grup: g.Nama_Grup, Nominal_Setoran: nominal, Periode: g.Periode,
    Jumlah_Slot: slot, Tanggal_Mulai: g.Tanggal_Mulai, Sistem_Undian: g.Sistem_Undian, Status: g.Status,
    putaranAktif: p, putaranTampil: pt, slot: slot, progress: Math.round(done / slot * 100),
    jumlahAnggota: ang.length, anggota: ang.slice(0, 4).map(function (a) { return a.Nama_Peserta; }),
    lunas: lunas, nextUndian: due ? fmtL_(due, 'dd/MM/yyyy') : '-', sisaHari: due ? selisihHari_(due, today) : 0,
    totalDana: nominal * ang.length
  };
}

function normSetoran_(s) {
  var o = {};
  Object.keys(s).forEach(function (k) { if (k !== '_i') o[k] = s[k]; });
  o.Nominal = num_(s.Nominal);
  o.Putaran_Ke = num_(s.Putaran_Ke);
  return o;
}

function normUndian_(u) {
  var o = {};
  Object.keys(u).forEach(function (k) { if (k !== '_i') o[k] = u[k]; });
  o.Total_Dana = num_(u.Total_Dana);
  o.Putaran_Ke = num_(u.Putaran_Ke);
  return o;
}

function anggotaAktif_() {
  return objs_(getTable_('Anggota_Grup')).filter(function (a) { return a.Status === 'Aktif'; });
}

// ==================== API: DASHBOARD ADMIN ====================
function getAdminDashboard(token) {
  return run_(function () {
    auth_(token, 'Admin');
    var key = 'DASH_' + dashVer_();
    var hit = cache_().get(key);
    if (hit) return JSON.parse(hit);
    var data = hitungDashboard_();
    try { cache_().put(key, JSON.stringify(data), 60); } catch (e) { /* data terlalu besar untuk cache */ }
    return data;
  });
}

function grafikBulan_(setoran, bulan, tahun) {
  var hari = new Date(tahun, bulan + 1, 0).getDate();
  var data = [], labels = [];
  for (var i = 1; i <= hari; i++) { data.push(0); labels.push(i); }
  var pb = bulan === 0 ? 11 : bulan - 1, pt = bulan === 0 ? tahun - 1 : tahun;
  var total = 0, prev = 0;
  setoran.forEach(function (s) {
    if (s.Status !== 'Lunas') return;
    var d = parseTgl_(s.Tanggal);
    if (!d) return;
    var n = num_(s.Nominal);
    if (d.getMonth() === bulan && d.getFullYear() === tahun) { data[d.getDate() - 1] += n; total += n; }
    else if (d.getMonth() === pb && d.getFullYear() === pt) prev += n;
  });
  var persen = prev ? Math.round((total - prev) / prev * 100) : (total ? 100 : 0);
  return { labels: labels, data: data, total: total, prevTotal: prev, persen: persen, bulan: bulan, tahun: tahun };
}

function getGrafikBulanan(token, bulan, tahun) {
  return run_(function () {
    auth_(token, 'Admin');
    return grafikBulan_(objs_(getTable_('Setoran')), num_(bulan), num_(tahun));
  });
}

function hitungDashboard_() {
  var today = hariIni_();
  var bln = today.getMonth(), thn = today.getFullYear();
  var users = objs_(getTable_('Users'));
  var grups = objs_(getTable_('Grup_Arisan'));
  var angg = anggotaAktif_();
  var setor = objs_(getTable_('Setoran'));
  var und = objs_(getTable_('Undian'));
  var tLog = getTable_('Log_Aktivitas');
  var smap = statusMap_(setor);
  var angByGrup = kelompok_(angg, 'ID_Grup');
  var fotoBy = fotoMap_(users);
  var gmap = {};
  grups.forEach(function (g) { gmap[g.ID_Grup] = g; });

  // --- Statistik utama ---
  var peserta = users.filter(function (u) { return u.Role === 'Peserta' && u.Status === 'Aktif'; });
  var pesertaBaru = peserta.filter(function (u) {
    var d = parseTgl_(u.Tanggal_Daftar);
    return d && d.getMonth() === bln && d.getFullYear() === thn;
  }).length;
  var grupAktif = grups.filter(function (g) { return g.Status === 'Aktif'; });
  var danaBulan = 0, danaHari = 0, menunggu = 0, menungguHari = 0, terkumpul = 0;
  var lunasList = [];
  setor.forEach(function (s) {
    var d = parseTgl_(s.Tanggal), n = num_(s.Nominal);
    if (s.Status === 'Lunas') {
      terkumpul += n;
      lunasList.push(s);
      if (d && d.getMonth() === bln && d.getFullYear() === thn) {
        danaBulan += n;
        if (d.getDate() === today.getDate()) danaHari += n;
      }
    } else if (s.Status === 'Menunggu Verifikasi') {
      menunggu++;
      if (d && selisihHari_(tglSaja_(d), today) === 0) menungguHari++;
    }
  });

  // --- Donut, jatuh tempo, kalender ---
  var donut = { lunas: 0, menunggu: 0, ditolak: 0, belum: 0, total: 0 };
  var tempo = [], events = [], kartu = [];
  var nextTempo = null;
  grupAktif.forEach(function (g) {
    var ang = angByGrup[g.ID_Grup] || [];
    var slot = num_(g.Jumlah_Slot) || 1;
    var p = Math.max(1, num_(g.Putaran_Aktif));
    var due = jatuhTempo_(g, p);
    var sisa = due ? selisihHari_(due, today) : 0;
    var belum = 0;
    ang.forEach(function (a) {
      var st = smap[g.ID_Grup + '|' + a.ID_User + '|' + p];
      donut.total++;
      if (st === 'Lunas') donut.lunas++;
      else if (st === 'Menunggu Verifikasi') donut.menunggu++;
      else {
        if (st === 'Ditolak') donut.ditolak++; else donut.belum++;
        belum++;
        tempo.push({
          nama: a.Nama_Peserta, idUser: a.ID_User, foto: fotoBy[a.ID_User] || '', grup: g.Nama_Grup, idGrup: g.ID_Grup, putaran: p,
          tanggal: due ? fmtL_(due, 'dd/MM/yyyy') : '-', sisa: sisa, nominal: num_(g.Nominal_Setoran), ditolak: st === 'Ditolak'
        });
      }
    });
    if (due && sisa >= 0 && (!nextTempo || sisa < nextTempo.sisa)) nextTempo = { sisa: sisa, teks: fmtL_(due, 'dd/MM/yyyy') + ' • ' + g.Nama_Grup };

    for (var k = 1; k <= slot; k++) {
      var d = jatuhTempo_(g, k);
      if (!d) continue;
      var key = fmtL_(d, 'yyyy-MM-dd');
      if (k < p) events.push({ tgl: key, tipe: 'selesai', judul: 'Undian putaran ' + k + ' selesai', grup: g.Nama_Grup });
      else if (k === p) {
        events.push({ tgl: key, tipe: 'tempo', judul: 'Jatuh tempo & undian putaran ' + k, grup: g.Nama_Grup });
        if (selisihHari_(d, today) < 0 && belum > 0) events.push({ tgl: key, tipe: 'telat', judul: belum + ' peserta belum setor (telat)', grup: g.Nama_Grup });
      } else events.push({ tgl: key, tipe: 'undian', judul: 'Jadwal undian putaran ' + k, grup: g.Nama_Grup });
    }
    kartu.push(grupCard_(g, ang, smap, today));
  });
  tempo.sort(function (a, b) { return a.sisa - b.sisa; });

  // --- Aktivitas terbaru (6 terakhir) ---
  var aktivitas = [];
  for (var i = tLog.rows.length - 1; i >= 0 && aktivitas.length < 6; i--) {
    if (String(tLog.rows[i][0]).trim() === '') continue;
    var o = rowObj_(tLog, i);
    aktivitas.push({ Timestamp: o.Timestamp, Nama: o.Nama, Role: o.Role, Aksi: o.Aksi, Detail: o.Detail, foto: fotoBy[o.ID_User] || '' });
  }

  // --- Top peserta disiplin (setor tepat waktu) ---
  var skor = {};
  lunasList.forEach(function (s) {
    var g = gmap[s.ID_Grup];
    if (!g) return;
    if (!skor[s.ID_User]) skor[s.ID_User] = { nama: s.Nama_Peserta, foto: fotoBy[s.ID_User] || '', tepat: 0, total: 0 };
    skor[s.ID_User].total++;
    var due = jatuhTempo_(g, num_(s.Putaran_Ke)), d = parseTgl_(s.Tanggal);
    if (due && d && selisihHari_(tglSaja_(d), due) <= 0) skor[s.ID_User].tepat++;
  });
  var top = Object.keys(skor).map(function (k) {
    var x = skor[k];
    x.persen = x.total ? Math.round(x.tepat / x.total * 100) : 0;
    return x;
  }).sort(function (a, b) { return b.tepat - a.tepat || b.persen - a.persen; }).slice(0, 5);

  // --- Keuangan ---
  var dicairkan = 0, belumCair = 0;
  und.forEach(function (u) {
    if (u.Status_Pencairan === 'Sudah Dicairkan') dicairkan += num_(u.Total_Dana);
    else belumCair += num_(u.Total_Dana);
  });
  var last = und.length ? normUndian_(und[und.length - 1]) : null;
  if (last) last.foto = fotoBy[last.ID_User] || '';

  return {
    stats: {
      peserta: peserta.length, pesertaBaru: pesertaBaru, grupAktif: grupAktif.length,
      grupSelesai: grups.length - grupAktif.length, menunggu: menunggu, menungguHari: menungguHari,
      danaBulan: danaBulan, danaHari: danaHari
    },
    grafik: grafikBulan_(setor, bln, thn),
    donut: donut,
    aktivitas: aktivitas,
    grup: kartu.slice(0, 4),
    events: events,
    tempo: tempo.slice(0, 8),
    pemenang: last,
    top: top,
    keuangan: { terkumpul: terkumpul, dicairkan: dicairkan, belumCair: belumCair, saldo: terkumpul - dicairkan },
    settings: getSettings_(),
    nextTempo: nextTempo ? nextTempo.teks : '-'
  };
}

// ==================== API: PENCARIAN GLOBAL ====================
function globalSearch(token, q) {
  return run_(function () {
    auth_(token, 'Admin');
    q = String(q || '').toLowerCase().trim();
    if (q.length < 2) return { peserta: [], setoran: [] };
    var peserta = objs_(getTable_('Users')).filter(function (u) {
      return u.Role === 'Peserta' && cocok_(q, [u.ID_User, u.Nama, u.No_HP]);
    }).slice(0, 5).map(function (u) { return { ID_User: u.ID_User, Nama: u.Nama, No_HP: u.No_HP }; });
    var setoran = objs_(getTable_('Setoran')).reverse().filter(function (s) {
      return cocok_(q, [s.ID_Transaksi, s.Nama_Peserta]);
    }).slice(0, 5).map(function (s) {
      return { ID_Transaksi: s.ID_Transaksi, Nama_Peserta: s.Nama_Peserta, Nominal: num_(s.Nominal), Status: s.Status };
    });
    return { peserta: peserta, setoran: setoran };
  });
}

// ==================== API: GRUP ARISAN ====================
function listGrup(token) {
  return run_(function () {
    auth_(token, 'Admin');
    var today = hariIni_();
    var grups = objs_(getTable_('Grup_Arisan'));
    var angBy = kelompok_(anggotaAktif_(), 'ID_Grup');
    var smap = statusMap_(objs_(getTable_('Setoran')));
    var out = grups.map(function (g) { return grupCard_(g, angBy[g.ID_Grup] || [], smap, today); });
    out.sort(function (a, b) { return (a.Status === 'Aktif' ? 0 : 1) - (b.Status === 'Aktif' ? 0 : 1); });
    return out;
  });
}

function saveGrup(token, d) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    d = d || {};
    var nama = clean_(d.nama, 80);
    var nominal = num_(d.nominal);
    var slot = num_(d.slot);
    var periode = d.periode === 'Mingguan' ? 'Mingguan' : 'Bulanan';
    var sistem = d.sistem === 'Urutan Tetap' ? 'Urutan Tetap' : 'Kocok Acak';
    var mulai = ymdKeDmy_(d.tanggalMulai);
    if (!nama) throw new Error('Nama grup wajib diisi.');
    if (nominal <= 0) throw new Error('Nominal setoran harus lebih dari 0.');
    if (slot < 2) throw new Error('Jumlah slot minimal 2.');
    if (!mulai) throw new Error('Tanggal mulai tidak valid.');

    return withLock_(function () {
      var t = getTable_('Grup_Arisan');
      var g;
      if (d.id) {
        var i = findIdx_(t, 'ID_Grup', d.id);
        if (i < 0) throw new Error('Grup tidak ditemukan.');
        var jml = anggotaAktif_().filter(function (a) { return a.ID_Grup === d.id; }).length;
        if (slot < jml) throw new Error('Slot tidak boleh kurang dari jumlah anggota (' + jml + ').');
        var status = d.status === 'Selesai' ? 'Selesai' : 'Aktif';
        g = updateObj_(t, i, { Nama_Grup: nama, Nominal_Setoran: nominal, Periode: periode, Jumlah_Slot: slot, Tanggal_Mulai: mulai, Sistem_Undian: sistem, Status: status });
        log_(s, 'Edit Grup', 'Grup ' + nama + ' diperbarui');
      } else {
        g = {
          ID_Grup: nextMasterId_(t, 'ID_Grup', 'GRP'), Nama_Grup: nama, Nominal_Setoran: nominal, Periode: periode,
          Jumlah_Slot: slot, Tanggal_Mulai: mulai, Sistem_Undian: sistem, Putaran_Aktif: 1, Status: 'Aktif', Dibuat_Pada: now_()
        };
        appendObjs_(t, [g]);
        log_(s, 'Buat Grup', 'Grup baru: ' + nama);
      }
      bump_();
      SpreadsheetApp.flush();
      var ang = anggotaAktif_().filter(function (a) { return a.ID_Grup === g.ID_Grup; });
      return grupCard_(g, ang, statusMap_(objs_(getTable_('Setoran'))), hariIni_());
    });
  });
}

function getGrupDetail(token, idGrup) {
  return run_(function () {
    auth_(token, 'Admin');
    var tG = getTable_('Grup_Arisan');
    var i = findIdx_(tG, 'ID_Grup', idGrup);
    if (i < 0) throw new Error('Grup tidak ditemukan.');
    var g = rowObj_(tG, i);
    var ang = anggotaAktif_().filter(function (a) { return a.ID_Grup === idGrup; })
      .sort(function (a, b) { return num_(a.No_Urut) - num_(b.No_Urut); });
    var setor = objs_(getTable_('Setoran')).filter(function (x) { return x.ID_Grup === idGrup; });
    var smap = statusMap_(setor);
    var slot = num_(g.Jumlah_Slot) || 1;
    var und = objs_(getTable_('Undian')).filter(function (u) { return u.ID_Grup === idGrup; });
    var pemenangBy = {};
    und.forEach(function (u) { pemenangBy[num_(u.Putaran_Ke)] = u.Nama_Pemenang; });
    var terkumpul = {};
    setor.forEach(function (x) { if (x.Status === 'Lunas') terkumpul[num_(x.Putaran_Ke)] = (terkumpul[num_(x.Putaran_Ke)] || 0) + num_(x.Nominal); });

    var anggota = ang.map(function (a) {
      var st = [];
      for (var k = 1; k <= slot; k++) st.push(smap[idGrup + '|' + a.ID_User + '|' + k] || '');
      return { ID_Anggota: a.ID_Anggota, ID_User: a.ID_User, Nama_Peserta: a.Nama_Peserta, No_Urut: num_(a.No_Urut), Status_Menang: a.Status_Menang, Putaran_Menang: num_(a.Putaran_Menang), status: st };
    });
    var putaran = [];
    for (var k = 1; k <= slot; k++) {
      var d = jatuhTempo_(g, k);
      putaran.push({ ke: k, tanggal: d ? fmtL_(d, 'dd/MM/yyyy') : '-', terkumpul: terkumpul[k] || 0, pemenang: pemenangBy[k] || '' });
    }
    return { grup: grupCard_(g, ang, smap, hariIni_()), anggota: anggota, putaran: putaran };
  });
}

function getAnggotaGrup(token, idGrup) {
  return run_(function () {
    auth_(token, 'Admin');
    var tG = getTable_('Grup_Arisan');
    var i = findIdx_(tG, 'ID_Grup', idGrup);
    if (i < 0) throw new Error('Grup tidak ditemukan.');
    var g = rowObj_(tG, i);
    var p = Math.max(1, num_(g.Putaran_Aktif));
    var smap = statusMap_(objs_(getTable_('Setoran')).filter(function (x) { return x.ID_Grup === idGrup; }));
    var ang = anggotaAktif_().filter(function (a) { return a.ID_Grup === idGrup; })
      .sort(function (a, b) { return num_(a.No_Urut) - num_(b.No_Urut); });
    return {
      grup: grupCard_(g, ang, smap, hariIni_()),
      anggota: ang.map(function (a) {
        return { ID_User: a.ID_User, Nama_Peserta: a.Nama_Peserta, No_Urut: num_(a.No_Urut), status: smap[idGrup + '|' + a.ID_User + '|' + p] || '' };
      })
    };
  });
}

// pairs: [{idGrup, idUser, nama}]
function tambahAnggota_(pairs) {
  var tA = getTable_('Anggota_Grup');
  var angg = objs_(tA);
  var gmap = {};
  objs_(getTable_('Grup_Arisan')).forEach(function (g) { gmap[g.ID_Grup] = g; });
  var maxId = 0;
  angg.forEach(function (a) { var m = String(a.ID_Anggota).match(/(\d+)$/); if (m) maxId = Math.max(maxId, +m[1]); });
  var baru = [], gagal = [];
  pairs.forEach(function (p) {
    var g = gmap[p.idGrup];
    if (!g) { gagal.push('Grup ' + p.idGrup + ' tidak ditemukan'); return; }
    if (g.Status !== 'Aktif') { gagal.push(g.Nama_Grup + ' sudah selesai'); return; }
    var list = angg.concat(baru).filter(function (a) { return a.ID_Grup === p.idGrup && a.Status === 'Aktif'; });
    if (list.some(function (a) { return a.ID_User === p.idUser; })) { gagal.push(p.nama + ' sudah ada di ' + g.Nama_Grup); return; }
    if (list.length >= num_(g.Jumlah_Slot)) { gagal.push(g.Nama_Grup + ' sudah penuh'); return; }
    var noUrut = list.reduce(function (mx, a) { return Math.max(mx, num_(a.No_Urut)); }, 0) + 1;
    maxId++;
    baru.push({
      ID_Anggota: 'AGT-' + pad4_(maxId), ID_Grup: p.idGrup, ID_User: p.idUser, Nama_Peserta: p.nama,
      No_Urut: noUrut, Status_Menang: 'Belum', Putaran_Menang: '', Tanggal_Gabung: now_(), Status: 'Aktif'
    });
  });
  appendObjs_(tA, baru);
  return { added: baru.length, gagal: gagal };
}

function addAnggotaGrup(token, idGrup, idUsers) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    if (!idUsers || !idUsers.length) throw new Error('Pilih minimal 1 peserta.');
    return withLock_(function () {
      var umap = {};
      objs_(getTable_('Users')).forEach(function (u) { umap[u.ID_User] = u; });
      var pairs = [];
      idUsers.forEach(function (id) {
        var u = umap[String(id)];
        if (u && u.Role === 'Peserta' && u.Status === 'Aktif') pairs.push({ idGrup: idGrup, idUser: u.ID_User, nama: u.Nama });
      });
      var r = tambahAnggota_(pairs);
      if (r.added) log_(s, 'Tambah Anggota', r.added + ' peserta ditambahkan ke grup ' + idGrup);
      bump_();
      SpreadsheetApp.flush();
      return r;
    });
  });
}

// ==================== API: PESERTA (ADMIN) ====================
function listPeserta(token, p) {
  return run_(function () {
    auth_(token, 'Admin');
    p = p || {};
    var q = String(p.q || '').toLowerCase().trim();
    var st = String(p.status || '');
    var grupBy = {};
    anggotaAktif_().forEach(function (a) { (grupBy[a.ID_User] = grupBy[a.ID_User] || []).push(a.ID_Grup); });
    var list = objs_(getTable_('Users')).filter(function (u) {
      return u.Role === 'Peserta' && (!st || u.Status === st) && cocok_(q, [u.ID_User, u.Nama, u.No_HP]);
    }).reverse().map(function (u) {
      var o = pubUser_(u);
      o.grupIds = grupBy[u.ID_User] || [];
      o.jumlahGrup = o.grupIds.length;
      return o;
    });
    return paginate_(list, p);
  });
}

function listPesertaRingkas(token) {
  return run_(function () {
    auth_(token, 'Admin');
    return objs_(getTable_('Users')).filter(function (u) { return u.Role === 'Peserta' && u.Status === 'Aktif'; })
      .map(function (u) { return { id: u.ID_User, nama: u.Nama, hp: u.No_HP }; })
      .sort(function (a, b) { return a.nama.localeCompare(b.nama); });
  });
}

function savePeserta(token, d) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    d = d || {};
    var nama = clean_(d.nama, 80);
    var hp = normHp_(d.hp);
    var alamat = clean_(d.alamat, 200);
    if (!nama) throw new Error('Nama peserta wajib diisi.');
    if (hp.length < 9 || hp.length > 15) throw new Error('No. HP tidak valid.');

    return withLock_(function () {
      var t = getTable_('Users');
      var semua = objs_(t);
      var dup = semua.some(function (u) { return normHp_(u.No_HP) === hp && u.ID_User !== d.id; });
      if (dup) throw new Error('No. HP sudah terdaftar.');
      var u;
      if (d.id) {
        var i = findIdx_(t, 'ID_User', d.id);
        if (i < 0) throw new Error('Peserta tidak ditemukan.');
        var lama = rowObj_(t, i);
        var status = d.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif';
        u = updateObj_(t, i, { Nama: nama, No_HP: hp, Alamat: alamat, Status: status });
        // Sinkron nama di Anggota_Grup (batch 1 kolom)
        if (lama.Nama !== nama) {
          var tA = getTable_('Anggota_Grup');
          var cU = tA.idx.ID_User, cN = tA.idx.Nama_Peserta;
          if (tA.rows.length && cU !== undefined && cN !== undefined) {
            var kol = tA.rows.map(function (r) { return [r[cU] === d.id ? nama : r[cN]]; });
            tA.sheet.getRange(2, cN + 1, kol.length, 1).setValues(kol);
          }
        }
        log_(s, 'Edit Peserta', 'Data ' + nama + ' diperbarui');
      } else {
        validPin_(d.pin);
        var salt = newSalt_();
        u = {
          ID_User: nextMasterId_(t, 'ID_User', 'USR'), Nama: nama, No_HP: hp, PIN_Hash: hashPin_(d.pin, salt), Salt: salt,
          Role: 'Peserta', Status: 'Aktif', Alamat: alamat, Tanggal_Daftar: now_(), Terakhir_Login: ''
        };
        appendObjs_(t, [u]);
        log_(s, 'Tambah Peserta', nama + ' terdaftar sebagai peserta');
      }
      var gagal = [];
      var ids = (d.grupIds || []).map(String);
      if (ids.length) {
        var r = tambahAnggota_(ids.map(function (g) { return { idGrup: g, idUser: u.ID_User, nama: nama }; }));
        gagal = r.gagal;
      }
      bump_();
      SpreadsheetApp.flush();
      var o = pubUser_(u);
      o.grupIds = anggotaAktif_().filter(function (a) { return a.ID_User === u.ID_User; }).map(function (a) { return a.ID_Grup; });
      o.jumlahGrup = o.grupIds.length;
      o.gagal = gagal.filter(function (x) { return x.indexOf('sudah ada') === -1; });
      return o;
    });
  });
}

function setStatusPeserta(token, idUser, status) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    status = status === 'Nonaktif' ? 'Nonaktif' : 'Aktif';
    return withLock_(function () {
      var t = getTable_('Users');
      var i = findIdx_(t, 'ID_User', idUser);
      if (i < 0) throw new Error('Peserta tidak ditemukan.');
      var u = updateObj_(t, i, { Status: status });
      log_(s, status === 'Aktif' ? 'Aktifkan Peserta' : 'Nonaktifkan Peserta', u.Nama + ' di-set ' + status);
      bump_();
      SpreadsheetApp.flush();
      return pubUser_(u);
    });
  });
}

function resetPin(token, idUser, pinBaru) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    validPin_(pinBaru);
    return withLock_(function () {
      var t = getTable_('Users');
      var i = findIdx_(t, 'ID_User', idUser);
      if (i < 0) throw new Error('Peserta tidak ditemukan.');
      var salt = newSalt_();
      var u = updateObj_(t, i, { PIN_Hash: hashPin_(pinBaru, salt), Salt: salt });
      cache_().remove('GAGAL_' + normHp_(u.No_HP));
      log_(s, 'Reset PIN', 'PIN ' + u.Nama + ' direset oleh admin');
      SpreadsheetApp.flush();
      return true;
    });
  });
}

// ==================== API: SETORAN ====================
function listSetoran(token, p) {
  return run_(function () {
    auth_(token, 'Admin');
    p = p || {};
    var q = String(p.q || '').toLowerCase().trim();
    var st = String(p.status || '');
    var g = String(p.idGrup || '');
    var base = objs_(getTable_('Setoran')).reverse().filter(function (s) {
      return (!g || s.ID_Grup === g) && cocok_(q, [s.ID_Transaksi, s.Nama_Peserta, s.Nama_Grup, s.ID_User]);
    });
    var counts = { semua: base.length, 'Menunggu Verifikasi': 0, 'Lunas': 0, 'Ditolak': 0 };
    base.forEach(function (s) { if (counts[s.Status] !== undefined) counts[s.Status]++; });
    var hasil = st ? base.filter(function (s) { return s.Status === st; }) : base;
    var out = paginate_(hasil, p);
    var fotoBy = fotoMap_();
    out.items = out.items.map(function (s) { var o = normSetoran_(s); o.foto = fotoBy[s.ID_User] || ''; return o; });
    out.counts = counts;
    return out;
  });
}

function buatSetoran_(t, g, u, putaran, metode, catatan, status, verifikator, struk) {
  var obj = {
    ID_Transaksi: nextDailyId_(t, 'ID_Transaksi', 'STR'), Tanggal: now_(), ID_Grup: g.ID_Grup, Nama_Grup: g.Nama_Grup,
    ID_User: u.id, Nama_Peserta: u.nama, Putaran_Ke: putaran, Nominal: num_(g.Nominal_Setoran),
    Metode: metode === 'Tunai' ? 'Tunai' : 'Transfer', Status: status, Catatan: clean_(catatan, 200),
    Diverifikasi_Oleh: verifikator || '', Tanggal_Verifikasi: verifikator ? now_() : '', Alasan_Tolak: '',
    Bukti_Struk: struk || '', No_Kwitansi: '', Kwitansi_PDF: ''
  };
  appendObjs_(t, [obj]);
  if (struk) gantiNama_(struk, obj.ID_Transaksi + '.jpg');
  return obj;
}

// Upload struk SEBELUM lock (agar antrean tulis tidak tertahan), hapus lagi jika proses gagal
function denganStruk_(dataUrl, prefix, fn) {
  var fileId = dataUrl ? simpanGambar_(dataUrl, prefix + '-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss'), FOLDER.struk) : '';
  try { return fn(fileId); } catch (e) { hapusFile_(fileId); throw e; }
}

function cekBolehSetor_(t, idGrup, idUser, putaran) {
  var ada = objs_(t).filter(function (x) {
    return x.ID_Grup === idGrup && x.ID_User === idUser && num_(x.Putaran_Ke) === putaran;
  });
  if (ada.some(function (x) { return x.Status === 'Lunas'; })) throw new Error('Setoran putaran ' + putaran + ' sudah lunas.');
  if (ada.some(function (x) { return x.Status === 'Menunggu Verifikasi'; })) throw new Error('Setoran putaran ' + putaran + ' masih menunggu verifikasi.');
}

function catatSetoranAdmin(token, d) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    d = d || {};
    return denganStruk_(d.struk, 'STRUK-' + clean_(d.idUser, 20), function (fileId) {
      return withLock_(function () {
        var tG = getTable_('Grup_Arisan');
        var gi = findIdx_(tG, 'ID_Grup', d.idGrup);
        if (gi < 0) throw new Error('Grup tidak ditemukan.');
        var g = rowObj_(tG, gi);
        var anggota = anggotaAktif_().filter(function (a) { return a.ID_Grup === g.ID_Grup && a.ID_User === d.idUser; })[0];
        if (!anggota) throw new Error('Peserta bukan anggota grup ini.');
        var putaran = num_(d.putaran) || Math.max(1, num_(g.Putaran_Aktif));
        if (putaran < 1 || putaran > num_(g.Jumlah_Slot)) throw new Error('Putaran tidak valid.');
        var t = getTable_('Setoran');
        cekBolehSetor_(t, g.ID_Grup, d.idUser, putaran);
        var obj = buatSetoran_(t, g, { id: anggota.ID_User, nama: anggota.Nama_Peserta }, putaran, d.metode, d.catatan, 'Lunas', s.nama, fileId);
        log_(s, 'Catat Setoran', 'Setoran ' + anggota.Nama_Peserta + ' Rp ' + num_(g.Nominal_Setoran) + ' dicatat (' + obj.ID_Transaksi + ')');
        bump_();
        SpreadsheetApp.flush();
        return normSetoran_(obj);
      });
    });
  });
}

function verifikasiSetoran(token, idTrx, aksi, alasan) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    if (aksi === 'Tolak' && !clean_(alasan)) throw new Error('Alasan penolakan wajib diisi.');
    return withLock_(function () {
      var t = getTable_('Setoran');
      var i = findIdx_(t, 'ID_Transaksi', idTrx);
      if (i < 0) throw new Error('Transaksi tidak ditemukan.');
      var cur = rowObj_(t, i);
      if (cur.Status !== 'Menunggu Verifikasi') throw new Error('Transaksi ini sudah diproses (' + cur.Status + ').');
      var o = updateObj_(t, i, {
        Status: aksi === 'Tolak' ? 'Ditolak' : 'Lunas', Diverifikasi_Oleh: s.nama,
        Tanggal_Verifikasi: now_(), Alasan_Tolak: aksi === 'Tolak' ? clean_(alasan, 200) : ''
      });
      log_(s, aksi === 'Tolak' ? 'Tolak Setoran' : 'Verifikasi Setoran', 'Setoran ' + o.Nama_Peserta + ' (' + idTrx + ') ' + (aksi === 'Tolak' ? 'ditolak' : 'diterima'));
      bump_();
      SpreadsheetApp.flush();
      return normSetoran_(o);
    });
  });
}

// ==================== API: GALERI BUKTI ====================
function listGaleri(token, p) {
  return run_(function () {
    auth_(token, 'Admin');
    p = p || {};
    var tG = getTable_('Grup_Arisan');
    var gi = findIdx_(tG, 'ID_Grup', p.idGrup);
    if (gi < 0) throw new Error('Pilih grup arisan dulu.');
    var g = rowObj_(tG, gi);
    var slot = num_(g.Jumlah_Slot) || 1;
    var put = num_(p.putaran) || Math.min(Math.max(1, num_(g.Putaran_Aktif)), slot);
    var st = String(p.status || '');
    var fotoBy = fotoMap_();
    var ang = anggotaAktif_().filter(function (a) { return a.ID_Grup === g.ID_Grup; });
    var urut = {};
    ang.forEach(function (a) { urut[a.ID_User] = num_(a.No_Urut); });

    var setor = objs_(getTable_('Setoran')).filter(function (x) { return x.ID_Grup === g.ID_Grup && num_(x.Putaran_Ke) === put; });
    var sudah = {};
    var items = setor.map(function (x) {
      sudah[x.ID_User] = 1;
      var o = normSetoran_(x);
      o.tipe = 'setoran';
      o.No_Urut = urut[x.ID_User] || 999;
      o.foto = fotoBy[x.ID_User] || '';
      return o;
    });
    // Anggota yang belum setor sama sekali di putaran ini
    ang.forEach(function (a) {
      if (sudah[a.ID_User]) return;
      items.push({ tipe: 'belum', ID_User: a.ID_User, Nama_Peserta: a.Nama_Peserta, No_Urut: num_(a.No_Urut), Status: 'Belum Setor', Nominal: num_(g.Nominal_Setoran), foto: fotoBy[a.ID_User] || '' });
    });
    var counts = { semua: items.length, 'Menunggu Verifikasi': 0, 'Lunas': 0, 'Ditolak': 0, 'Belum Setor': 0, tanpaStruk: 0 };
    items.forEach(function (x) {
      if (counts[x.Status] !== undefined) counts[x.Status]++;
      if (x.tipe === 'setoran' && !x.Bukti_Struk) counts.tanpaStruk++;
    });
    items.sort(function (a, b) { return a.No_Urut - b.No_Urut || (a.Tanggal < b.Tanggal ? -1 : 1); });
    var hasil = st ? items.filter(function (x) { return x.Status === st; }) : items;
    var out = paginate_(hasil, p);
    out.counts = counts;
    out.putaran = put;
    out.slot = slot;
    return out;
  });
}

// ==================== API: GAMBAR & FOTO PROFIL ====================
function uploadFotoProfil(token, dataUrl) {
  return run_(function () {
    var s = auth_(token);
    var fileId = simpanGambar_(dataUrl, 'FOTO-' + s.id + '-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss'), FOLDER.foto);
    try {
      return withLock_(function () {
        var t = getTable_('Users');
        var i = findIdx_(t, 'ID_User', s.id);
        if (i < 0) throw new Error('User tidak ditemukan.');
        var lama = rowObj_(t, i).Foto_Profil;
        updateObj_(t, i, { Foto_Profil: fileId });
        updateSesi_(token, { foto: fileId });
        log_(s, 'Ganti Foto', s.nama + ' memperbarui foto profil');
        bump_();
        SpreadsheetApp.flush();
        hapusFile_(lama);
        return { foto: fileId };
      });
    } catch (e) { hapusFile_(fileId); throw e; }
  });
}

// Fallback tampilan gambar via backend (base64) jika link publik diblokir
function getGambar(token, fileId) {
  return run_(function () {
    var s = auth_(token);
    fileId = String(fileId || '');
    if (!/^[A-Za-z0-9_-]{10,}$/.test(fileId)) throw new Error('ID gambar tidak valid.');
    var file = DriveApp.getFileById(fileId);
    var jenis = null, ids = folderIds_(), par = file.getParents();
    while (par.hasNext()) { var pid = par.next().getId(); if (ids[pid]) jenis = ids[pid]; }
    if (!jenis) throw new Error('Akses gambar ditolak.');
    // Peserta hanya boleh melihat struk/kwitansi miliknya sendiri
    if (s.role !== 'Admin' && jenis !== 'foto') {
      var milik = objs_(getTable_('Setoran')).some(function (x) { return x.ID_User === s.id && (x.Bukti_Struk === fileId || x.Kwitansi_PDF === fileId); });
      if (!milik) throw new Error('Akses gambar ditolak.');
    }
    var blob = file.getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  });
}

// ==================== API: KWITANSI PDF ====================
function terbilang_(n) {
  n = Math.floor(Math.abs(n));
  var s = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
  if (n < 12) return s[n];
  if (n < 20) return terbilang_(n - 10) + ' belas';
  if (n < 100) return terbilang_(Math.floor(n / 10)) + ' puluh' + (n % 10 ? ' ' + terbilang_(n % 10) : '');
  if (n < 200) return 'seratus' + (n - 100 ? ' ' + terbilang_(n - 100) : '');
  if (n < 1000) return terbilang_(Math.floor(n / 100)) + ' ratus' + (n % 100 ? ' ' + terbilang_(n % 100) : '');
  if (n < 2000) return 'seribu' + (n - 1000 ? ' ' + terbilang_(n - 1000) : '');
  if (n < 1e6) return terbilang_(Math.floor(n / 1000)) + ' ribu' + (n % 1000 ? ' ' + terbilang_(n % 1000) : '');
  if (n < 1e9) return terbilang_(Math.floor(n / 1e6)) + ' juta' + (n % 1e6 ? ' ' + terbilang_(n % 1e6) : '');
  return terbilang_(Math.floor(n / 1e9)) + ' miliar' + (n % 1e9 ? ' ' + terbilang_(n % 1e9) : '');
}
function rupiah_(n) { return 'Rp ' + String(num_(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function escHtml_(v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function htmlKwitansi_(x, set, strukUri) {
  var perusahaan = escHtml_(set.Nama_Perusahaan || 'PT. Berjaya Makmur Indonesia');
  var nominal = num_(x.Nominal);
  var tb = terbilang_(nominal);
  tb = tb ? tb.charAt(0).toUpperCase() + tb.slice(1) + ' rupiah' : 'Nol rupiah';
  var baris = function (k, v) { return '<tr><td style="padding:7px 0;color:#6B7280;width:38%">' + k + '</td><td style="padding:7px 0;font-weight:bold;color:#111827">' + v + '</td></tr>'; };
  return '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,Helvetica,sans-serif;margin:0;padding:28px;color:#111827;font-size:12px">' +
    '<div style="border:2px solid #1F7A4D;border-radius:12px;overflow:hidden">' +
    '<div style="background:#1F7A4D;color:#ffffff;padding:18px 22px">' +
    '<div style="font-size:20px;font-weight:bold">' + perusahaan + '</div>' +
    '<div style="font-size:11px;opacity:.85;margin-top:2px">' + escHtml_(set.Nama_Aplikasi || 'Arisan Berjaya') + (set.Kontak_Admin ? ' &bull; Kontak: ' + escHtml_(set.Kontak_Admin) : '') + '</div></div>' +
    '<div style="padding:20px 22px">' +
    '<table style="width:100%;border-collapse:collapse"><tr>' +
    '<td><div style="font-size:22px;font-weight:bold;letter-spacing:2px;color:#1F7A4D">KWITANSI</div><div style="color:#6B7280;margin-top:2px">Bukti Pembayaran Setoran Arisan</div></td>' +
    '<td style="text-align:right;vertical-align:top"><div style="color:#6B7280">No. Kwitansi</div><div style="font-size:14px;font-weight:bold">' + escHtml_(x.No_Kwitansi) + '</div></td>' +
    '</tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:16px;border-top:1px solid #E3E8E5">' +
    baris('Telah terima dari', escHtml_(x.Nama_Peserta) + ' (' + escHtml_(x.ID_User) + ')') +
    baris('Untuk pembayaran', 'Setoran ' + escHtml_(x.Nama_Grup) + ' &bull; Putaran ke-' + num_(x.Putaran_Ke)) +
    baris('ID Transaksi', escHtml_(x.ID_Transaksi)) +
    baris('Metode', escHtml_(x.Metode)) +
    baris('Tanggal setor', escHtml_(x.Tanggal)) +
    baris('Tanggal verifikasi', escHtml_(x.Tanggal_Verifikasi || '-')) +
    '</table>' +
    '<div style="margin-top:14px;background:#E9F7EF;border-radius:10px;padding:14px 16px">' +
    '<div style="color:#6B7280;font-size:11px">Jumlah</div><div style="font-size:24px;font-weight:bold;color:#1F7A4D">' + rupiah_(nominal) + '</div>' +
    '<div style="font-style:italic;color:#374151;margin-top:4px">Terbilang: ' + escHtml_(tb) + '</div></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:18px"><tr>' +
    '<td style="vertical-align:top;width:55%">' + (strukUri ? '<div style="color:#6B7280;margin-bottom:6px">Lampiran bukti struk:</div><img src="' + strukUri + '" style="max-width:220px;max-height:220px;border:1px solid #E3E8E5;border-radius:8px">' : '<div style="color:#9CA3AF">Tanpa lampiran struk (' + escHtml_(x.Metode) + ')</div>') + '</td>' +
    '<td style="vertical-align:bottom;text-align:center">' +
    '<div style="display:inline-block;border:2px solid #2E9E5B;color:#2E9E5B;border-radius:8px;padding:6px 14px;font-weight:bold;letter-spacing:2px">LUNAS</div>' +
    '<div style="margin-top:28px;color:#6B7280">Diverifikasi oleh</div><div style="font-weight:bold;margin-top:4px">' + escHtml_(x.Diverifikasi_Oleh || '-') + '</div></td>' +
    '</tr></table></div>' +
    '<div style="background:#F3F6F4;color:#6B7280;font-size:10px;padding:10px 22px">Dokumen ini dibuat otomatis oleh sistem ' + escHtml_(set.Nama_Aplikasi || 'Arisan Berjaya') + ' dan sah tanpa tanda tangan basah. Dicetak: ' + now_() + '</div>' +
    '</div></body></html>';
}

// Buat kwitansi (idempoten): jika sudah ada & file masih ada, kembalikan yang lama
function kwitansi_(idTrx) {
  // 1) Reservasi nomor kwitansi di dalam lock
  var x = withLock_(function () {
    var t = getTable_('Setoran');
    var i = findIdx_(t, 'ID_Transaksi', idTrx);
    if (i < 0) throw new Error('Transaksi tidak ditemukan.');
    var o = rowObj_(t, i);
    if (o.Status !== 'Lunas') throw new Error('Kwitansi hanya untuk setoran yang sudah lunas.');
    if (!o.No_Kwitansi) {
      o = updateObj_(t, i, { No_Kwitansi: nextDailyId_(t, 'No_Kwitansi', 'KWT') });
      SpreadsheetApp.flush();
    }
    return o;
  });
  if (x.Kwitansi_PDF) {
    try { var lama = DriveApp.getFileById(x.Kwitansi_PDF); if (!lama.isTrashed()) return { id: lama.getId(), no: x.No_Kwitansi, file: lama }; } catch (e) { /* file hilang, buat ulang */ }
  }
  // 2) Render PDF di luar lock (proses berat)
  var strukUri = '';
  if (x.Bukti_Struk) {
    try { var b = DriveApp.getFileById(x.Bukti_Struk).getBlob(); strukUri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) { strukUri = ''; }
  }
  var pdf = Utilities.newBlob(htmlKwitansi_(x, getSettings_(), strukUri), 'text/html', 'kwitansi.html').getAs('application/pdf');
  pdf.setName(x.No_Kwitansi + ' - ' + x.Nama_Peserta + '.pdf');
  var file = folder_(FOLDER.kwitansi).createFile(pdf);
  bagikan_(file);
  // 3) Simpan File ID (baca ulang tabel agar tidak menimpa perubahan lain)
  withLock_(function () {
    var t = getTable_('Setoran');
    var i = findIdx_(t, 'ID_Transaksi', idTrx);
    if (i > -1) updateObj_(t, i, { Kwitansi_PDF: file.getId() });
    SpreadsheetApp.flush();
  });
  return { id: file.getId(), no: x.No_Kwitansi, file: file };
}

function buatKwitansi(token, idTrx) {
  return run_(function () {
    auth_(token, 'Admin');
    var k = kwitansi_(idTrx);
    return { id: k.id, no: k.no };
  });
}

function unduhKwitansi(token, idTrx) {
  return run_(function () {
    var s = auth_(token);
    if (s.role !== 'Admin') {
      var t = getTable_('Setoran');
      var i = findIdx_(t, 'ID_Transaksi', idTrx);
      if (i < 0 || rowObj_(t, i).ID_User !== s.id) throw new Error('Kwitansi tidak ditemukan.');
    }
    var k = kwitansi_(idTrx);
    var blob = k.file.getBlob();
    return {
      id: k.id, no: k.no, nama: k.file.getName(), url: 'https://drive.google.com/file/d/' + k.id + '/view',
      data: 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes())
    };
  });
}

// ==================== API: UNDIAN ====================
function infoUndian_(idGrup) {
  var tG = getTable_('Grup_Arisan');
  var gi = findIdx_(tG, 'ID_Grup', idGrup);
  if (gi < 0) throw new Error('Grup tidak ditemukan.');
  var g = rowObj_(tG, gi);
  var p = Math.max(1, num_(g.Putaran_Aktif));
  var ang = anggotaAktif_().filter(function (a) { return a.ID_Grup === idGrup; });
  var smap = statusMap_(objs_(getTable_('Setoran')).filter(function (x) { return x.ID_Grup === idGrup; }));
  var sudah = objs_(getTable_('Undian')).some(function (u) { return u.ID_Grup === idGrup && num_(u.Putaran_Ke) === p; });
  var lunasCount = 0;
  var kandidat = ang.filter(function (a) { return a.Status_Menang !== 'Sudah'; }).map(function (a) {
    var lunas = smap[idGrup + '|' + a.ID_User + '|' + p] === 'Lunas';
    return { ID_User: a.ID_User, Nama_Peserta: a.Nama_Peserta, No_Urut: num_(a.No_Urut), lunas: lunas, _i: a._i };
  }).sort(function (a, b) { return a.No_Urut - b.No_Urut; });
  ang.forEach(function (a) { if (smap[idGrup + '|' + a.ID_User + '|' + p] === 'Lunas') lunasCount++; });
  var due = jatuhTempo_(g, Math.min(p, num_(g.Jumlah_Slot)));
  return {
    g: g, tG: tG, gi: gi, putaran: p, sudahDiundi: sudah, kandidat: kandidat, jumlahAnggota: ang.length,
    lunasCount: lunasCount, totalDana: num_(g.Nominal_Setoran) * ang.length, jatuhTempo: due ? fmtL_(due, 'dd/MM/yyyy') : '-',
    card: grupCard_(g, ang, smap, hariIni_())
  };
}

function getUndianInfo(token, idGrup) {
  return run_(function () {
    auth_(token, 'Admin');
    var x = infoUndian_(idGrup);
    return {
      grup: x.card, putaran: x.putaran, sudahDiundi: x.sudahDiundi, jumlahAnggota: x.jumlahAnggota,
      lunasCount: x.lunasCount, totalDana: x.totalDana, jatuhTempo: x.jatuhTempo,
      kandidat: x.kandidat.map(function (k) { return { ID_User: k.ID_User, Nama_Peserta: k.Nama_Peserta, No_Urut: k.No_Urut, lunas: k.lunas }; })
    };
  });
}

function lakukanUndian(token, idGrup, hanyaLunas) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    return withLock_(function () {
      var x = infoUndian_(idGrup);
      var g = x.g;
      if (g.Status !== 'Aktif') throw new Error('Grup sudah selesai.');
      if (x.sudahDiundi) throw new Error('Putaran ' + x.putaran + ' sudah diundi.');
      var pool = hanyaLunas ? x.kandidat.filter(function (k) { return k.lunas; }) : x.kandidat;
      if (!pool.length) throw new Error(hanyaLunas ? 'Belum ada kandidat yang lunas di putaran ini.' : 'Tidak ada kandidat tersisa.');

      var pemenang = g.Sistem_Undian === 'Urutan Tetap' ? pool[0] : pool[Math.floor(Math.random() * pool.length)];

      var tU = getTable_('Undian');
      var und = {
        ID_Undian: nextDailyId_(tU, 'ID_Undian', 'UND'), Tanggal: now_(), ID_Grup: g.ID_Grup, Nama_Grup: g.Nama_Grup,
        Putaran_Ke: x.putaran, ID_User: pemenang.ID_User, Nama_Pemenang: pemenang.Nama_Peserta, Total_Dana: x.totalDana,
        Status_Pencairan: 'Belum Dicairkan', Tanggal_Pencairan: '', Dilakukan_Oleh: s.nama
      };
      appendObjs_(tU, [und]);

      var tA = getTable_('Anggota_Grup');
      updateObj_(tA, pemenang._i, { Status_Menang: 'Sudah', Putaran_Menang: x.putaran });

      var slot = num_(g.Jumlah_Slot);
      var next = x.putaran + 1;
      updateObj_(x.tG, x.gi, { Putaran_Aktif: Math.min(next, slot), Status: next > slot ? 'Selesai' : 'Aktif' });

      log_(s, 'Undian', 'Pemenang putaran ' + x.putaran + ' ' + g.Nama_Grup + ': ' + pemenang.Nama_Peserta);
      bump_();
      SpreadsheetApp.flush();
      return {
        pemenang: { ID_User: pemenang.ID_User, Nama: pemenang.Nama_Peserta }, putaran: x.putaran, totalDana: x.totalDana,
        kandidat: pool.map(function (k) { return k.Nama_Peserta; }), undian: normUndian_(und), selesai: next > slot
      };
    });
  });
}

function listUndian(token, p) {
  return run_(function () {
    auth_(token, 'Admin');
    p = p || {};
    var g = String(p.idGrup || '');
    var fotoBy = fotoMap_();
    var list = objs_(getTable_('Undian')).reverse().filter(function (u) { return !g || u.ID_Grup === g; })
      .map(function (u) { var o = normUndian_(u); o.foto = fotoBy[u.ID_User] || ''; return o; });
    return paginate_(list, p);
  });
}

function setPencairan(token, idUndian) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    return withLock_(function () {
      var t = getTable_('Undian');
      var i = findIdx_(t, 'ID_Undian', idUndian);
      if (i < 0) throw new Error('Data undian tidak ditemukan.');
      var o = updateObj_(t, i, { Status_Pencairan: 'Sudah Dicairkan', Tanggal_Pencairan: now_() });
      log_(s, 'Pencairan Dana', 'Dana ' + o.Nama_Pemenang + ' (' + o.Nama_Grup + ') dicairkan');
      bump_();
      SpreadsheetApp.flush();
      return normUndian_(o);
    });
  });
}

// ==================== API: LAPORAN ====================
function filterLaporan_(p) {
  p = p || {};
  var g = String(p.idGrup || ''), st = String(p.status || '');
  var dari = p.dari ? parseTgl_(ymdKeDmy_(p.dari)) : null;
  var sampai = p.sampai ? parseTgl_(ymdKeDmy_(p.sampai)) : null;
  var dalam = function (s) {
    var d = tglSaja_(parseTgl_(s));
    if (!d) return !dari && !sampai;
    return (!dari || d >= dari) && (!sampai || d <= sampai);
  };
  var setor = objs_(getTable_('Setoran')).filter(function (s) {
    return (!g || s.ID_Grup === g) && (!st || s.Status === st) && dalam(s.Tanggal);
  });
  var und = objs_(getTable_('Undian')).filter(function (u) {
    return (!g || u.ID_Grup === g) && u.Status_Pencairan === 'Sudah Dicairkan' && dalam(u.Tanggal_Pencairan);
  });
  return { setor: setor, und: und };
}

function getLaporan(token, p) {
  return run_(function () {
    auth_(token, 'Admin');
    p = p || {};
    var f = filterLaporan_(p);
    var sum = { jumlah: f.setor.length, lunas: 0, menunggu: 0, ditolak: 0, nominalLunas: 0, nominalMenunggu: 0, dicairkan: 0 };
    var per = {};
    objs_(getTable_('Grup_Arisan')).forEach(function (g) {
      if (!p.idGrup || g.ID_Grup === p.idGrup) per[g.ID_Grup] = { nama: g.Nama_Grup, terkumpul: 0, dicairkan: 0, saldo: 0 };
    });
    f.setor.forEach(function (s) {
      var n = num_(s.Nominal);
      if (s.Status === 'Lunas') { sum.lunas++; sum.nominalLunas += n; if (per[s.ID_Grup]) per[s.ID_Grup].terkumpul += n; }
      else if (s.Status === 'Menunggu Verifikasi') { sum.menunggu++; sum.nominalMenunggu += n; }
      else if (s.Status === 'Ditolak') sum.ditolak++;
    });
    f.und.forEach(function (u) {
      sum.dicairkan += num_(u.Total_Dana);
      if (per[u.ID_Grup]) per[u.ID_Grup].dicairkan += num_(u.Total_Dana);
    });
    var perGrup = Object.keys(per).map(function (k) { var x = per[k]; x.saldo = x.terkumpul - x.dicairkan; return x; });
    var out = paginate_(f.setor.reverse().map(normSetoran_), p);
    out.summary = sum;
    out.perGrup = perGrup;
    return out;
  });
}

function exportLaporan(token, p) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    var f = filterLaporan_(p);
    var nama = 'Laporan Arisan Berjaya - ' + Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HHmm');
    var file = SpreadsheetApp.create(nama);
    var sh = file.getSheets()[0];
    sh.setName('Setoran');
    var head = ['ID Transaksi', 'Tanggal', 'Grup', 'Peserta', 'Putaran', 'Nominal', 'Metode', 'Status', 'Diverifikasi Oleh', 'Catatan'];
    var rows = [head].concat(f.setor.map(function (x) {
      return [x.ID_Transaksi, x.Tanggal, x.Nama_Grup, x.Nama_Peserta, num_(x.Putaran_Ke), num_(x.Nominal), x.Metode, x.Status, x.Diverifikasi_Oleh, x.Catatan];
    }));
    sh.getRange(1, 1, rows.length, head.length).setNumberFormat('@').setValues(rows.map(function (r) { return r.map(String); }));
    if (rows.length > 1) sh.getRange(2, 6, rows.length - 1, 1).setNumberFormat('#,##0').setValues(rows.slice(1).map(function (r) { return [r[5]]; }));
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#1F7A4D').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, head.length);
    log_(s, 'Export Laporan', 'Export ' + f.setor.length + ' transaksi');
    SpreadsheetApp.flush();
    return { url: file.getUrl(), nama: nama };
  });
}

// ==================== API: PENGATURAN & LOG ====================
function getPengaturan(token) {
  return run_(function () { auth_(token); return getSettings_(); });
}

function savePengaturan(token, data) {
  return run_(function () {
    var s = auth_(token, 'Admin');
    data = data || {};
    return withLock_(function () {
      var t = getTable_('Pengaturan');
      var cK = t.idx.Key, cV = t.idx.Value;
      var baru = [];
      KUNCI_PENGATURAN.forEach(function (k) {
        if (data[k] === undefined) return;
        var v = clean_(data[k], 300);
        var i = findIdx_(t, 'Key', k);
        if (i > -1) t.rows[i][cV] = v; else baru.push({ Key: k, Value: v });
      });
      if (t.rows.length) {
        t.sheet.getRange(2, cV + 1, t.rows.length, 1).setNumberFormat('@').setValues(t.rows.map(function (r) { return [r[cV]]; }));
      }
      appendObjs_(t, baru);
      log_(s, 'Pengaturan', 'Pengaturan aplikasi diperbarui');
      bump_();
      SpreadsheetApp.flush();
      return getSettings_();
    });
  });
}

function listLog(token, p) {
  return run_(function () {
    auth_(token, 'Admin');
    var fotoBy = fotoMap_();
    var list = objs_(getTable_('Log_Aktivitas')).reverse().map(function (o) {
      return { Timestamp: o.Timestamp, Nama: o.Nama, Role: o.Role, Aksi: o.Aksi, Detail: o.Detail, foto: fotoBy[o.ID_User] || '' };
    });
    return paginate_(list, p);
  });
}

// ==================== API: PESERTA ====================
function getPesertaDashboard(token) {
  return run_(function () {
    var s = auth_(token, 'Peserta');
    var today = hariIni_();
    var tU = getTable_('Users');
    var ui = findIdx_(tU, 'ID_User', s.id);
    if (ui < 0) throw new Error('SESI_HABIS');
    var me = pubUser_(rowObj_(tU, ui));
    var semuaAng = anggotaAktif_();
    var milikku = semuaAng.filter(function (a) { return a.ID_User === s.id; });
    var angBy = kelompok_(semuaAng, 'ID_Grup');
    var gmap = {};
    objs_(getTable_('Grup_Arisan')).forEach(function (g) { gmap[g.ID_Grup] = g; });
    var setor = objs_(getTable_('Setoran'));
    var smap = statusMap_(setor);
    var totalSetoran = 0, jumlahLunas = 0;
    setor.forEach(function (x) { if (x.ID_User === s.id && x.Status === 'Lunas') { totalSetoran += num_(x.Nominal); jumlahLunas++; } });

    var grup = milikku.map(function (a) {
      var g = gmap[a.ID_Grup];
      if (!g) return null;
      var c = grupCard_(g, angBy[g.ID_Grup] || [], smap, today);
      var st = smap[g.ID_Grup + '|' + s.id + '|' + c.putaranAktif] || 'Belum Bayar';
      c.noUrut = num_(a.No_Urut);
      c.menang = a.Status_Menang === 'Sudah';
      c.putaranMenang = num_(a.Putaran_Menang);
      c.statusSaya = c.Status === 'Selesai' ? 'Selesai' : st;
      c.bisaSetor = c.Status === 'Aktif' && (st === 'Belum Bayar' || st === 'Ditolak');
      return c;
    }).filter(function (x) { return x; });

    var idGrupku = milikku.map(function (a) { return a.ID_Grup; });
    var fotoBy = fotoMap_();
    var pemenang = objs_(getTable_('Undian')).filter(function (u) { return idGrupku.indexOf(u.ID_Grup) > -1; })
      .reverse().slice(0, 5).map(function (u) { var o = normUndian_(u); o.foto = fotoBy[u.ID_User] || ''; return o; });

    return { user: me, grup: grup, pemenang: pemenang, totalSetoran: totalSetoran, jumlahLunas: jumlahLunas, settings: getSettings_() };
  });
}

function kirimSetoran(token, d) {
  return run_(function () {
    var s = auth_(token, 'Peserta');
    d = d || {};
    if (d.metode !== 'Tunai' && !d.struk) throw new Error('Bukti struk wajib dilampirkan untuk metode Transfer.');
    return denganStruk_(d.struk, 'STRUK-' + s.id, function (fileId) {
      return withLock_(function () {
        var tG = getTable_('Grup_Arisan');
        var gi = findIdx_(tG, 'ID_Grup', d.idGrup);
        if (gi < 0) throw new Error('Grup tidak ditemukan.');
        var g = rowObj_(tG, gi);
        if (g.Status !== 'Aktif') throw new Error('Grup ini sudah selesai.');
        var anggota = anggotaAktif_().filter(function (a) { return a.ID_Grup === g.ID_Grup && a.ID_User === s.id; })[0];
        if (!anggota) throw new Error('Kamu bukan anggota grup ini.');
        var putaran = Math.max(1, num_(g.Putaran_Aktif));
        var t = getTable_('Setoran');
        cekBolehSetor_(t, g.ID_Grup, s.id, putaran);
        var obj = buatSetoran_(t, g, { id: s.id, nama: s.nama }, putaran, d.metode, d.catatan, 'Menunggu Verifikasi', '', fileId);
        log_(s, 'Kirim Setoran', s.nama + ' mengirim setoran Rp ' + num_(g.Nominal_Setoran) + ' (' + g.Nama_Grup + ')' + (fileId ? ' + struk' : ''));
        bump_();
        SpreadsheetApp.flush();
        return normSetoran_(obj);
      });
    });
  });
}

function listSetoranSaya(token, p) {
  return run_(function () {
    var s = auth_(token, 'Peserta');
    var list = objs_(getTable_('Setoran')).filter(function (x) { return x.ID_User === s.id; }).reverse().map(normSetoran_);
    return paginate_(list, p);
  });
}
