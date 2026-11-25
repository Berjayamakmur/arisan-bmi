/**
 * ============================================================
 *  ARISAN BERJAYA — setup.gs (Inisialisasi Database / Safe Migrate)
 *  PT. Berjaya Makmur Indonesia
 * ------------------------------------------------------------
 *  Jalankan fungsi setupDatabase() dari editor Apps Script.
 *  AMAN dijalankan berulang kali:
 *   - Sheet yang belum ada akan dibuat.
 *   - Header yang kurang akan ditambahkan di ujung kanan.
 *   - TIDAK ADA clear() / hapus baris data pengguna.
 *   - Data dummy hanya diisi saat database benar-benar kosong.
 *   - Folder Google Drive (Foto Profil, Struk Setoran, Kwitansi) dibuat jika belum ada.
 *  Catatan: memakai helper hashPin_(), newSalt_(), pad4_(), folder_(), FOLDER & KOLOM_ANGKA dari code.gs
 * ============================================================
 */

var SCHEMA = {
  Users: ['ID_User', 'Nama', 'No_HP', 'PIN_Hash', 'Salt', 'Role', 'Status', 'Alamat', 'Tanggal_Daftar', 'Terakhir_Login', 'Foto_Profil'],
  Grup_Arisan: ['ID_Grup', 'Nama_Grup', 'Nominal_Setoran', 'Periode', 'Jumlah_Slot', 'Tanggal_Mulai', 'Sistem_Undian', 'Putaran_Aktif', 'Status', 'Dibuat_Pada'],
  Anggota_Grup: ['ID_Anggota', 'ID_Grup', 'ID_User', 'Nama_Peserta', 'No_Urut', 'Status_Menang', 'Putaran_Menang', 'Tanggal_Gabung', 'Status'],
  Setoran: ['ID_Transaksi', 'Tanggal', 'ID_Grup', 'Nama_Grup', 'ID_User', 'Nama_Peserta', 'Putaran_Ke', 'Nominal', 'Metode', 'Status', 'Catatan', 'Diverifikasi_Oleh', 'Tanggal_Verifikasi', 'Alasan_Tolak', 'Bukti_Struk', 'No_Kwitansi', 'Kwitansi_PDF'],
  Undian: ['ID_Undian', 'Tanggal', 'ID_Grup', 'Nama_Grup', 'Putaran_Ke', 'ID_User', 'Nama_Pemenang', 'Total_Dana', 'Status_Pencairan', 'Tanggal_Pencairan', 'Dilakukan_Oleh'],
  Log_Aktivitas: ['Timestamp', 'ID_User', 'Nama', 'Role', 'Aksi', 'Detail'],
  Pengaturan: ['Key', 'Value']
};

// Menu custom di Spreadsheet
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🌱 Arisan Berjaya')
      .addItem('Setup / Migrasi Database', 'setupDatabase')
      .addToUi();
  } catch (e) { /* dijalankan tanpa UI */ }
}

// ==================== FUNGSI UTAMA ====================
function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var laporan = [];
    Object.keys(SCHEMA).forEach(function (nama) {
      laporan.push('• ' + nama + ': ' + migrasiSheet_(ss, nama, SCHEMA[nama]));
    });
    SpreadsheetApp.flush();
    isiDummy_(ss);
    SpreadsheetApp.flush();
    laporan.push('• Folder Drive: ' + siapkanFolder_());
    Logger.log('✅ Setup selesai\n' + laporan.join('\n'));
  } catch (e) {
    Logger.log('❌ Setup gagal: ' + e.message);
    throw e;
  }
}

// ==================== SAFE MIGRATE PER SHEET ====================
function migrasiSheet_(ss, nama, headers) {
  var sh = ss.getSheetByName(nama);
  var baru = false;
  if (!sh) { sh = ss.insertSheet(nama); baru = true; }

  var lastCol = sh.getLastColumn();
  var existing = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] : [];
  var adaHeader = existing.some(function (h) { return String(h).trim() !== ''; });
  var kolomBaru = [];
  var finalHeader;

  if (!adaHeader) {
    // Header kosong: jika ternyata ada data di bawahnya, sisipkan baris baru (data TIDAK disentuh)
    if (sh.getLastRow() > 0) sh.insertRowBefore(1);
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    finalHeader = headers.slice();
    kolomBaru = headers.map(function (h, i) { return i + 1; });
  } else {
    // Header sudah ada: hanya tambahkan kolom yang belum ada di ujung kanan
    finalHeader = existing.map(function (h) { return String(h).trim(); });
    headers.forEach(function (h) {
      if (finalHeader.indexOf(h) === -1) { finalHeader.push(h); kolomBaru.push(finalHeader.length); }
    });
    if (kolomBaru.length) {
      if (sh.getMaxColumns() < finalHeader.length) sh.insertColumnsAfter(sh.getMaxColumns(), finalHeader.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);
    }
  }

  // Format kolom baru saja: angka -> '0', selain itu -> teks '@' (agar No HP & tanggal tidak berubah)
  var maxRow = sh.getMaxRows();
  if (maxRow > 1) {
    kolomBaru.forEach(function (c) {
      var h = finalHeader[c - 1];
      sh.getRange(2, c, maxRow - 1, 1).setNumberFormat(KOLOM_ANGKA.indexOf(h) > -1 ? '0' : '@');
    });
  }

  // Styling header
  sh.getRange(1, 1, 1, finalHeader.length)
    .setFontWeight('bold').setBackground('#1F7A4D').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);

  if (baru) return 'dibuat baru';
  return kolomBaru.length ? 'ditambah ' + kolomBaru.length + ' kolom' : 'aman, tidak ada perubahan';
}

// Buat folder Drive untuk foto profil, struk & kwitansi (aman dijalankan berulang, folder lama dipakai ulang)
function siapkanFolder_() {
  var nama = Object.keys(FOLDER).map(function (k) { return folder_(FOLDER[k]).getName(); });
  return FOLDER_ROOT + ' / ' + nama.join(', ');
}

// Tulis array object sesuai urutan header aktual sheet (1x setValues)
function tulisObj_(sh, objs) {
  if (!objs.length) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  var rows = objs.map(function (o) {
    return headers.map(function (h) { return o[h] !== undefined && o[h] !== null ? o[h] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

// ==================== DATA DUMMY ====================
function isiDummy_(ss) {
  var tz = Session.getScriptTimeZone();
  var f = function (d, p) { return Utilities.formatDate(d, tz, p || 'dd/MM/yyyy HH:mm:ss'); };
  var base = new Date();
  var hari = function (offset, jam, menit) {
    return new Date(base.getFullYear(), base.getMonth(), base.getDate() + offset, jam || 9, menit || 0, 0);
  };

  // --- Pengaturan default (diisi jika kosong) ---
  var shSet = ss.getSheetByName('Pengaturan');
  if (shSet.getLastRow() <= 1) {
    tulisObj_(shSet, [
      { Key: 'Nama_Aplikasi', Value: 'Arisan Berjaya' },
      { Key: 'Nama_Perusahaan', Value: 'PT. Berjaya Makmur Indonesia' },
      { Key: 'Kontak_Admin', Value: '0812-0000-0000' },
      { Key: 'Nama_Bank', Value: 'Bank BCA' },
      { Key: 'No_Rekening', Value: '1234567890' },
      { Key: 'Atas_Nama', Value: 'PT. Berjaya Makmur Indonesia' },
      { Key: 'Quote', Value: 'Sedikit demi sedikit, lama-lama jadi berjaya.' }
    ]);
  }

  // --- Dummy transaksi hanya jika SEMUA sheet utama kosong ---
  var cek = ['Users', 'Grup_Arisan', 'Anggota_Grup', 'Setoran', 'Undian', 'Log_Aktivitas'];
  var semuaKosong = cek.every(function (n) { return ss.getSheetByName(n).getLastRow() <= 1; });
  if (!semuaKosong) {
    Logger.log('ℹ️ Data sudah ada — data dummy dilewati (Safe Migrate).');
    return;
  }

  // Users: index 0 = admin, 1..10 = peserta
  var namaPeserta = ['Siti Rahmawati', 'Budi Santoso', 'Dewi Lestari', 'Agus Prasetyo', 'Rina Marlina',
    'Joko Susilo', 'Lina Kurniawati', 'Hendra Wijaya', 'Maya Sari', 'Fajar Nugroho'];
  var users = [];
  var saltA = newSalt_();
  users.push({
    ID_User: 'USR-0001', Nama: 'Admin Berjaya', No_HP: '081200000000', PIN_Hash: hashPin_('123456', saltA), Salt: saltA,
    Role: 'Admin', Status: 'Aktif', Alamat: 'Kantor Pusat PT. Berjaya Makmur Indonesia', Tanggal_Daftar: f(hari(-60)), Terakhir_Login: ''
  });
  namaPeserta.forEach(function (nm, i) {
    var s = newSalt_();
    users.push({
      ID_User: 'USR-' + pad4_(i + 2), Nama: nm, No_HP: '0812111100' + ('0' + (i + 1)).slice(-2),
      PIN_Hash: hashPin_('111111', s), Salt: s, Role: 'Peserta', Status: 'Aktif', Alamat: 'Jakarta',
      Tanggal_Daftar: f(hari(i < 8 ? -45 : -4)), Terakhir_Login: ''
    });
  });
  tulisObj_(ss.getSheetByName('Users'), users);

  // Grup arisan
  var g1Mulai = hari(-25), g2Mulai = hari(-9);
  var G1 = {
    ID_Grup: 'GRP-0001', Nama_Grup: 'Arisan Bulanan Berjaya', Nominal_Setoran: 500000, Periode: 'Bulanan', Jumlah_Slot: 10,
    Tanggal_Mulai: f(g1Mulai, 'dd/MM/yyyy'), Sistem_Undian: 'Kocok Acak', Putaran_Aktif: 2, Status: 'Aktif', Dibuat_Pada: f(hari(-30))
  };
  var G2 = {
    ID_Grup: 'GRP-0002', Nama_Grup: 'Arisan Mingguan Makmur', Nominal_Setoran: 100000, Periode: 'Mingguan', Jumlah_Slot: 6,
    Tanggal_Mulai: f(g2Mulai, 'dd/MM/yyyy'), Sistem_Undian: 'Urutan Tetap', Putaran_Aktif: 2, Status: 'Aktif', Dibuat_Pada: f(hari(-12))
  };
  tulisObj_(ss.getSheetByName('Grup_Arisan'), [G1, G2]);

  // Anggota grup
  var anggota = [], no = 0, k;
  for (k = 1; k <= 10; k++) {
    anggota.push({
      ID_Anggota: 'AGT-' + pad4_(++no), ID_Grup: G1.ID_Grup, ID_User: users[k].ID_User, Nama_Peserta: users[k].Nama, No_Urut: k,
      Status_Menang: k === 3 ? 'Sudah' : 'Belum', Putaran_Menang: k === 3 ? 1 : '', Tanggal_Gabung: f(hari(-26)), Status: 'Aktif'
    });
  }
  for (k = 1; k <= 6; k++) {
    anggota.push({
      ID_Anggota: 'AGT-' + pad4_(++no), ID_Grup: G2.ID_Grup, ID_User: users[k].ID_User, Nama_Peserta: users[k].Nama, No_Urut: k,
      Status_Menang: k === 1 ? 'Sudah' : 'Belum', Putaran_Menang: k === 1 ? 1 : '', Tanggal_Gabung: f(hari(-10)), Status: 'Aktif'
    });
  }
  tulisObj_(ss.getSheetByName('Anggota_Grup'), anggota);

  // Setoran
  var st = [];
  var add = function (d, g, idx, putaran, status, metode, catatan, alasan) {
    var u = users[idx];
    var verif = status !== 'Menunggu Verifikasi';
    st.push({
      _d: d, Tanggal: f(d), ID_Grup: g.ID_Grup, Nama_Grup: g.Nama_Grup, ID_User: u.ID_User, Nama_Peserta: u.Nama,
      Putaran_Ke: putaran, Nominal: g.Nominal_Setoran, Metode: metode, Status: status, Catatan: catatan || '',
      Diverifikasi_Oleh: verif ? 'Admin Berjaya' : '', Tanggal_Verifikasi: verif ? f(new Date(d.getTime() + 3600000)) : '',
      Alasan_Tolak: alasan || ''
    });
  };
  for (k = 1; k <= 10; k++) add(hari(-28 + (k % 3), 8 + k, 15), G1, k, 1, 'Lunas', k % 2 ? 'Transfer' : 'Tunai');
  [1, 2, 4, 5].forEach(function (x, i) { add(hari(-3 + (i % 3), 10 + i), G1, x, 2, 'Lunas', 'Transfer'); });
  add(hari(0, 8, 30), G1, 6, 2, 'Menunggu Verifikasi', 'Transfer', 'Transfer via BCA');
  add(hari(0, 9, 45), G1, 7, 2, 'Menunggu Verifikasi', 'Transfer');
  add(hari(-2, 14), G1, 8, 2, 'Ditolak', 'Transfer', '', 'Bukti transfer tidak sesuai nominal');
  for (k = 1; k <= 6; k++) add(hari(-10 + (k % 2), 9 + k), G2, k, 1, 'Lunas', 'Tunai');
  add(hari(-3, 11), G2, 2, 2, 'Lunas', 'Tunai');
  add(hari(-3, 13), G2, 3, 2, 'Lunas', 'Transfer');
  add(hari(0, 7, 50), G2, 4, 2, 'Menunggu Verifikasi', 'Transfer');

  // Urutkan kronologis lalu buat ID sekuensial reset harian
  st.sort(function (a, b) { return a._d - b._d; });
  var ctr = {};
  st.forEach(function (s) {
    var tag = f(s._d, 'ddMMyy');
    ctr[tag] = (ctr[tag] || 0) + 1;
    s.ID_Transaksi = 'STR-' + tag + '-' + pad4_(ctr[tag]);
    delete s._d;
  });
  tulisObj_(ss.getSheetByName('Setoran'), st);

  // Undian
  tulisObj_(ss.getSheetByName('Undian'), [
    {
      ID_Undian: 'UND-' + f(hari(-25), 'ddMMyy') + '-0001', Tanggal: f(hari(-25, 15)), ID_Grup: G1.ID_Grup, Nama_Grup: G1.Nama_Grup,
      Putaran_Ke: 1, ID_User: users[3].ID_User, Nama_Pemenang: users[3].Nama, Total_Dana: 5000000,
      Status_Pencairan: 'Sudah Dicairkan', Tanggal_Pencairan: f(hari(-24, 10)), Dilakukan_Oleh: 'Admin Berjaya'
    },
    {
      ID_Undian: 'UND-' + f(hari(-9), 'ddMMyy') + '-0001', Tanggal: f(hari(-9, 16)), ID_Grup: G2.ID_Grup, Nama_Grup: G2.Nama_Grup,
      Putaran_Ke: 1, ID_User: users[1].ID_User, Nama_Pemenang: users[1].Nama, Total_Dana: 600000,
      Status_Pencairan: 'Belum Dicairkan', Tanggal_Pencairan: '', Dilakukan_Oleh: 'Admin Berjaya'
    }
  ]);

  // Log aktivitas
  var adm = users[0];
  var lg = function (d, u, aksi, detail) {
    return { Timestamp: f(d), ID_User: u.ID_User, Nama: u.Nama, Role: u.Role, Aksi: aksi, Detail: detail };
  };
  tulisObj_(ss.getSheetByName('Log_Aktivitas'), [
    lg(hari(-25, 15), adm, 'Undian', 'Pemenang putaran 1 Arisan Bulanan Berjaya: Dewi Lestari'),
    lg(hari(-9, 16), adm, 'Undian', 'Pemenang putaran 1 Arisan Mingguan Makmur: Siti Rahmawati'),
    lg(hari(-4, 10), adm, 'Tambah Peserta', 'Maya Sari & Fajar Nugroho terdaftar sebagai peserta'),
    lg(hari(-2, 14, 30), adm, 'Tolak Setoran', 'Setoran Hendra Wijaya ditolak'),
    lg(hari(0, 7, 50), users[4], 'Kirim Setoran', 'Agus Prasetyo mengirim setoran Rp 100000 (Arisan Mingguan Makmur)'),
    lg(hari(0, 8, 30), users[6], 'Kirim Setoran', 'Joko Susilo mengirim setoran Rp 500000 (Arisan Bulanan Berjaya)'),
    lg(hari(0, 9, 45), users[7], 'Kirim Setoran', 'Lina Kurniawati mengirim setoran Rp 500000 (Arisan Bulanan Berjaya)')
  ]);

  Logger.log('🌱 Data dummy dibuat.\n' +
    '👤 Admin   : No HP 081200000000 | PIN 123456 (SEGERA GANTI setelah login!)\n' +
    '👥 Peserta : No HP 081211110001 s/d 081211110010 | PIN 111111');
}
