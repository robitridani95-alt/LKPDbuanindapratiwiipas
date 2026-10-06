/**
 * LKPD IPAS SMK - Backend Google Apps Script (v2: bank soal + admin)
 * Spreadsheet = database. Jalankan setup() sekali, lalu Deploy > Web app.
 * Jika sebelumnya sudah memakai versi lama: tempel kode ini, jalankan setup() lagi,
 * lalu Deploy > Manage deployments > Edit > New version.
 */
var SH_SISWA = 'Siswa', SH_HASIL = 'Hasil', SH_SET = 'Pengaturan', SH_SOAL = 'Soal', SH_CAT = 'Catatan', SH_UP = 'CatatanSiswa';
var MAX_GAGAL = 5;      // batas salah PIN / kata sandi
var KUNCI_MENIT = 10;   // lama blokir (menit)

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Jalankan SEKALI (dan lagi setelah update kode) untuk membuat sheet & header. */
function setup() {
  var ss = ss_();
  var s = ss.getSheetByName(SH_SISWA) || ss.insertSheet(SH_SISWA);
  s.getRange('A:D').setNumberFormat('@');
  if (s.getLastRow() === 0) {
    s.appendRow(['NIS', 'Nama', 'Kelas', 'PIN']);
    s.appendRow(['2024001', 'Contoh Siswa', 'X-TKJ', '1234']);
  }
  s.getRange('A1:D1').setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  s.setFrozenRows(1);

  var h = ss.getSheetByName(SH_HASIL) || ss.insertSheet(SH_HASIL);
  h.getRange('A1:P1').setValues([['NIS', 'Nama', 'Kelas', 'T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7',
    'Total', 'Nilai', 'Status', 'Jawaban (JSON)', 'Diperbarui', 'Jumlah soal']])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  h.getRange('A:A').setNumberFormat('@');
  h.setFrozenRows(1);

  var q = ss.getSheetByName(SH_SOAL) || ss.insertSheet(SH_SOAL);
  q.getRange('A1:L1').setValues([['ID', 'Topik (1-7)', 'Soal', 'Opsi A', 'Opsi B', 'Opsi C', 'Opsi D',
    'Kunci (A-D)', 'Pembahasan', 'Model 3D (0-21)', 'Gambar (ID Drive)', 'Perintah siswa']])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  q.setFrozenRows(1);

  var c = sheetCat_();
  c.getRange('A1:D1').setValues([['ID', 'Topik (1-7)', 'Perintah catatan', 'Gambar (ID Drive)']])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  c.setFrozenRows(1);

  var u = sheetUp_();
  u.getRange('A1:H1').setValues([['NIS', 'Nama', 'Kelas', 'ID catatan', 'Topik', 'Foto catatan', 'Waktu', 'ID file']])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
  u.getRange('A:A').setNumberFormat('@');
  u.setFrozenRows(1);

  var p = ss.getSheetByName(SH_SET) || ss.insertSheet(SH_SET);
  if (!p.getRange('B1').getValue()) {
    p.getRange('A1').setValue('KKM');
    p.getRange('B1').setValue(75).setBackground('#ffff99');
  }
  p.getRange('A2').setValue('Jumlah soal');
  p.getRange('B2').setFormula('=COUNTA(Soal!A2:A)');
  p.getRange('A3').setValue('Kata sandi guru');
  if (!p.getRange('B3').getValue()) p.getRange('B3').setValue('ganti-saya').setBackground('#ffff99');
  p.getRange('A5').setValue('Ubah KKM di B1 dan kata sandi guru di B3 (WAJIB diganti). Soal dikelola dari halaman LKPD > Masuk sebagai guru.');
  p.setColumnWidth(1, 160);
  SpreadsheetApp.flush();
}

function doGet() { return json_({ ok: true, msg: 'API LKPD IPAS aktif.' }); }

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (String(d.a || '').indexOf('adm_') === 0) return json_(admin_(d));

    var nis = String(d.nis || '').replace(/[^A-Za-z0-9_-]/g, '');
    var pin = String(d.pin || '').trim();
    if (!nis || !pin) return json_({ ok: false, msg: 'Isi NIS dan PIN.' });

    if (d.a === 'daftar') {
      var rg = daftar_(nis, d.nama, pin);
      if (!rg.ok) return json_(rg);
      d.a = 'login'; // lanjut login otomatis
    }

    var cache = CacheService.getScriptCache(), kKey = 'gagal_' + nis;
    if (Number(cache.get(kKey) || 0) >= MAX_GAGAL)
      return json_({ ok: false, msg: 'Terlalu banyak percobaan salah. Coba lagi ' + KUNCI_MENIT + ' menit lagi.' });

    var siswa = cari_(nis);
    if (!siswa || siswa.pin !== pin) {
      cache.put(kKey, String(Number(cache.get(kKey) || 0) + 1), KUNCI_MENIT * 60);
      return json_({ ok: false, msg: 'NIS atau PIN salah. Hubungi guru jika lupa.' });
    }
    cache.remove(kKey);

    if (d.a === 'login')
      return json_({ ok: true, nama: siswa.nama, kelas: siswa.kelas, st: ambilState_(nis), soal: daftarSoal_(), cat: daftarCatatan_(), up: uploadSiswa_(nis) });
    if (d.a === 'upload') return json_(unggahCatatan_(siswa, d.cid, d.img));
    if (d.a === 'save') {
      simpan_(siswa, d.sc || {}, String(d.st || '').slice(0, 40000));
      return json_({ ok: true });
    }
    return json_({ ok: false, msg: 'Aksi tidak dikenal.' });
  } catch (err) {
    return json_({ ok: false, msg: 'Kesalahan server: ' + err.message });
  }
}

/* ---------- Pendaftaran mandiri siswa ---------- */
function daftar_(nis, nama, pin) {
  nama = String(nama || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  if (nama.length < 3) return { ok: false, msg: 'Isi nama lengkap (minimal 3 huruf).' };
  if (!/^\d{4,8}$/.test(pin)) return { ok: false, msg: 'PIN harus 4-8 angka.' };
  if (nis.length < 3 || nis.length > 20) return { ok: false, msg: 'NIS tidak valid.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (cari_(nis)) return { ok: false, msg: 'NIS sudah terdaftar. Silakan masuk, atau hubungi guru jika lupa PIN.' };
    var s = ss_().getSheetByName(SH_SISWA), r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 4).setNumberFormat('@').setValues([[nis, nama, '', pin]]);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- Admin: kelola bank soal ---------- */
function admin_(d) {
  var c = CacheService.getScriptCache(), k = 'gagal_adm';
  if (Number(c.get(k) || 0) >= MAX_GAGAL) return { ok: false, msg: 'Terlalu banyak percobaan. Coba lagi ' + KUNCI_MENIT + ' menit lagi.' };
  var pw = String(ss_().getSheetByName(SH_SET).getRange('B3').getDisplayValue()).trim();
  if (!pw || String(d.key || '') !== pw) {
    c.put(k, String(Number(c.get(k) || 0) + 1), KUNCI_MENIT * 60);
    return { ok: false, msg: 'Kata sandi guru salah.' };
  }
  c.remove(k);
  if (d.a === 'adm_list') return { ok: true, items: daftarSoal_() };
  if (d.a === 'adm_clist') return { ok: true, cat: daftarCatatan_() };
  if (d.a === 'adm_hasil') return dataKumpul_();
  if (d.a === 'adm_foto') return fotoSiswa_(d.fid);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.a === 'adm_save') return simpanSoal_(d.item);
    if (d.a === 'adm_del') return hapusSoal_(d.id);
    if (d.a === 'adm_clear') return kosongkanSoal_();
    if (d.a === 'adm_csave') return simpanCatatan_(d.item);
    if (d.a === 'adm_cdel') return hapusCatatan_(d.id);
  } finally { lock.releaseLock(); }
  return { ok: false, msg: 'Aksi tidak dikenal.' };
}

function sheetSoal_() { return ss_().getSheetByName(SH_SOAL); }

function daftarSoal_() {
  var sh = sheetSoal_(), n = sh.getLastRow();
  if (n < 2) return [];
  var v = sh.getRange(2, 1, n - 1, 12).getDisplayValues(), out = [];
  for (var i = 0; i < v.length; i++) {
    if (!v[i][0]) continue;
    out.push({ id: v[i][0], t: Number(v[i][1]), s: v[i][2], o: [v[i][3], v[i][4], v[i][5], v[i][6]],
      k: 'ABCD'.indexOf(String(v[i][7]).toUpperCase()), e: v[i][8], m: Number(v[i][9]) || 0,
      g: v[i][10], p: v[i][11], r: i });
  }
  out.sort(function (a, b) { return a.t - b.t || a.r - b.r; });
  return out;
}

function bersih_(it) {
  var o = (it.o || []).map(function (x) { return String(x || '').trim().slice(0, 200); });
  var t = parseInt(it.t, 10), k = parseInt(it.k, 10), m = parseInt(it.m, 10) || 0;
  var s = String(it.s || '').trim().slice(0, 500);
  if (!(t >= 1 && t <= 7)) throw new Error('Topik harus 1-7.');
  if (!s) throw new Error('Soal tidak boleh kosong.');
  if (o.length !== 4 || o.some(function (x) { return !x; })) throw new Error('Isi keempat opsi jawaban.');
  if (!(k >= 0 && k <= 3)) throw new Error('Pilih kunci jawaban.');
  if (m < 0 || m > 21) m = 0;
  return [t, s, o[0], o[1], o[2], o[3], 'ABCD'.charAt(k), String(it.e || '').trim().slice(0, 500), m,
    '', String(it.p || '').trim().slice(0, 1000)];
}

function barisSoal_(id) {
  var sh = sheetSoal_(), n = sh.getLastRow();
  if (n < 2 || !id) return 0;
  var v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === String(id)) return i + 2;
  return 0;
}

/* ---------- Gambar soal (disimpan di Google Drive) ---------- */
function folderGambar_() {
  var it = DriveApp.getFoldersByName('LKPD IPAS Gambar');
  return it.hasNext() ? it.next() : DriveApp.createFolder('LKPD IPAS Gambar');
}

function simpanGambar_(dataUrl) {
  var m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('Format gambar tidak valid.');
  if (m[2].length > 4000000) throw new Error('Gambar terlalu besar.');
  var ext = m[1].split('/')[1];
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'soal-' + Date.now() + '.' + ext);
  var f = folderGambar_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}

function hapusGambar_(id) {
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
}

/* ---------- Unggahan catatan siswa (foto JPG, privat di Drive guru) ---------- */
function sheetUp_() { return ss_().getSheetByName(SH_UP) || ss_().insertSheet(SH_UP); }

function uploadSiswa_(nis) {
  var sh = sheetUp_(), n = sh.getLastRow(), out = {};
  if (n < 2) return out;
  var v = sh.getRange(2, 1, n - 1, 7).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === nis) out[v[i][3]] = v[i][6];
  return out;
}

function unggahCatatan_(siswa, cid, img) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var r = barisCat_(cid);
    if (!r) return { ok: false, msg: 'Catatan tidak ditemukan.' };
    var m = /^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/.exec(String(img || ''));
    if (!m) return { ok: false, msg: 'Format harus JPG.' };
    if (m[1].length > 4000000) return { ok: false, msg: 'Foto terlalu besar.' };
    var topik = sheetCat_().getRange(r, 2).getValue();
    var it = DriveApp.getFoldersByName('LKPD IPAS Catatan Siswa');
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder('LKPD IPAS Catatan Siswa');
    var nama = (siswa.nis + '_' + siswa.nama + '_T' + topik + '_' + cid).replace(/[^\w\-]+/g, '_') + '.jpg';
    var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', nama));
    var sh = sheetUp_(), n = sh.getLastRow(), row = 0, lama = '';
    if (n > 1) {
      var v = sh.getRange(2, 1, n - 1, 8).getDisplayValues();
      for (var i = 0; i < v.length; i++)
        if (String(v[i][0]).trim() === siswa.nis && v[i][3] === String(cid)) { row = i + 2; lama = v[i][7]; break; }
    }
    if (!row) row = n + 1;
    var waktu = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
    sh.getRange(row, 1, 1, 5).setValues([[siswa.nis, siswa.nama, siswa.kelas, String(cid), topik]]);
    sh.getRange(row, 6).setFormula('=HYPERLINK("https://drive.google.com/file/d/' + f.getId() + '/view","Buka foto")');
    sh.getRange(row, 7, 1, 2).setValues([[waktu, f.getId()]]);
    hapusGambar_(lama);
    return { ok: true, w: waktu };
  } finally { lock.releaseLock(); }
}

/* ---------- Admin: pengumpulan siswa ---------- */
function dataKumpul_() {
  var ss = ss_(), s = ss.getSheetByName(SH_SISWA), n = s.getLastRow(), siswa = [];
  if (n > 1) s.getRange(2, 1, n - 1, 3).getDisplayValues().forEach(function (v) {
    if (String(v[0]).trim()) siswa.push({ nis: String(v[0]).trim(), nama: v[1], kelas: v[2] });
  });
  var h = ss.getSheetByName(SH_HASIL), m = h.getLastRow(), hasil = {};
  if (m > 1) h.getRange(2, 1, m - 1, 16).getDisplayValues().forEach(function (v) {
    var ref = [];
    try { var st = JSON.parse(v[13] || '{}'); for (var i = 0; i < 7; i++) ref.push(String(st['r' + i] || '')); }
    catch (e) { ref = ['', '', '', '', '', '', '']; }
    hasil[String(v[0]).trim()] = { sc: v.slice(3, 10), total: v[10], nilai: v[11], status: v[12], w: v[14], n: v[15], ref: ref };
  });
  var u = sheetUp_(), k = u.getLastRow(), up = [];
  if (k > 1) u.getRange(2, 1, k - 1, 8).getDisplayValues().forEach(function (v) {
    if (v[0]) up.push({ nis: String(v[0]).trim(), cid: v[3], w: v[6], fid: v[7] });
  });
  return { ok: true, siswa: siswa, hasil: hasil, up: up, cat: daftarCatatan_() };
}

function fotoSiswa_(fid) {
  fid = String(fid || '');
  var u = sheetUp_(), k = u.getLastRow();
  if (k < 2 || !fid) return { ok: false, msg: 'Foto tidak ditemukan.' };
  var ids = u.getRange(2, 8, k - 1, 1).getDisplayValues(), ada = false;
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === fid) { ada = true; break; }
  if (!ada) return { ok: false, msg: 'Foto tidak ditemukan.' };
  try {
    var b = DriveApp.getFileById(fid).getBlob();
    return { ok: true, img: 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes()) };
  } catch (e) { return { ok: false, msg: 'File foto tidak bisa dibuka.' }; }
}

/* ---------- Catatan bergambar (perintah guru + gambar) ---------- */
function sheetCat_() { return ss_().getSheetByName(SH_CAT) || ss_().insertSheet(SH_CAT); }

function daftarCatatan_() {
  var sh = sheetCat_(), n = sh.getLastRow();
  if (n < 2) return [];
  var v = sh.getRange(2, 1, n - 1, 4).getDisplayValues(), out = [];
  for (var i = 0; i < v.length; i++) {
    if (!v[i][0]) continue;
    out.push({ id: v[i][0], t: Number(v[i][1]), p: v[i][2], g: v[i][3], r: i });
  }
  out.sort(function (a, b) { return a.t - b.t || a.r - b.r; });
  return out;
}

function barisCat_(id) {
  var sh = sheetCat_(), n = sh.getLastRow();
  if (n < 2 || !id) return 0;
  var v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === String(id)) return i + 2;
  return 0;
}

function simpanCatatan_(it) {
  try {
    var t = parseInt(it.t, 10), p = String(it.p || '').trim().slice(0, 1000);
    if (!(t >= 1 && t <= 7)) throw new Error('Topik harus 1-7.');
    if (!p) throw new Error('Perintah catatan tidak boleh kosong.');
    var sh = sheetCat_(), r = barisCat_(it.id);
    var lama = r ? String(sh.getRange(r, 4).getDisplayValue()) : '', g = lama;
    if (it.img) { g = simpanGambar_(it.img); hapusGambar_(lama); }
    else if (it.gdel) { g = ''; hapusGambar_(lama); }
    if (r) sh.getRange(r, 2, 1, 3).setValues([[t, p, g]]);
    else sh.appendRow([Utilities.getUuid().slice(0, 8), t, p, g]);
    return { ok: true, cat: daftarCatatan_() };
  } catch (err) { return { ok: false, msg: err.message }; }
}

function hapusCatatan_(id) {
  var r = barisCat_(id);
  if (!r) return { ok: false, msg: 'Catatan tidak ditemukan.' };
  hapusGambar_(String(sheetCat_().getRange(r, 4).getDisplayValue()));
  sheetCat_().deleteRow(r);
  return { ok: true, cat: daftarCatatan_() };
}

function simpanSoal_(it) {
  try {
    var row = bersih_(it), sh = sheetSoal_(), r = barisSoal_(it.id);
    var lama = r ? String(sh.getRange(r, 11).getDisplayValue()) : '';
    if (it.img) { row[9] = simpanGambar_(it.img); hapusGambar_(lama); }
    else if (it.gdel) { row[9] = ''; hapusGambar_(lama); }
    else row[9] = lama;
    if (r) sh.getRange(r, 2, 1, 11).setValues([row]);
    else sh.appendRow([Utilities.getUuid().slice(0, 8)].concat(row));
    return { ok: true, items: daftarSoal_() };
  } catch (err) { return { ok: false, msg: err.message }; }
}

function hapusSoal_(id) {
  var r = barisSoal_(id);
  if (!r) return { ok: false, msg: 'Soal tidak ditemukan.' };
  hapusGambar_(String(sheetSoal_().getRange(r, 11).getDisplayValue()));
  sheetSoal_().deleteRow(r);
  return { ok: true, items: daftarSoal_() };
}

function kosongkanSoal_() {
  var sh = sheetSoal_(), n = sh.getLastRow();
  if (n > 1) {
    sh.getRange(2, 11, n - 1, 1).getDisplayValues().forEach(function (x) { hapusGambar_(x[0]); });
    sh.deleteRows(2, n - 1);
  }
  return { ok: true, items: [] };
}

/* ---------- Siswa & hasil ---------- */
function cari_(nis) {
  var s = ss_().getSheetByName(SH_SISWA), n = s.getLastRow();
  if (n < 2) return null;
  var v = s.getRange(2, 1, n - 1, 4).getDisplayValues();
  for (var i = 0; i < v.length; i++)
    if (String(v[i][0]).trim() === nis)
      return { nis: nis, nama: v[i][1], kelas: v[i][2], pin: String(v[i][3]).trim() };
  return null;
}

function barisHasil_(nis) {
  var h = ss_().getSheetByName(SH_HASIL), n = h.getLastRow();
  if (n < 2) return 0;
  var v = h.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === nis) return i + 2;
  return 0;
}

function ambilState_(nis) {
  var r = barisHasil_(nis);
  return r ? String(ss_().getSheetByName(SH_HASIL).getRange(r, 14).getValue() || '') : '';
}

function simpan_(siswa, sc, st) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = ss_().getSheetByName(SH_HASIL);
    var r = barisHasil_(siswa.nis) || h.getLastRow() + 1;
    var skor = [];
    for (var i = 0; i < 7; i++) {
      var x = sc[i];
      skor.push(x === undefined || x === null || x === '' ? '' : Math.max(0, Math.min(99, parseInt(x, 10) || 0)));
    }
    h.getRange(r, 1, 1, 3).setValues([[siswa.nis, siswa.nama, siswa.kelas]]);
    h.getRange(r, 4, 1, 7).setValues([skor]);
    h.getRange(r, 11, 1, 3).setFormulas([[
      '=IF(COUNT(D' + r + ':J' + r + ')=0,"",SUM(D' + r + ':J' + r + '))',
      '=IF(OR(K' + r + '="",N(P' + r + ')=0),"",ROUND(K' + r + '/P' + r + '*100,0))',
      '=IF(L' + r + '="","",IF(L' + r + '>=Pengaturan!$B$1,"Tuntas","Remedial"))'
    ]]);
    h.getRange(r, 14, 1, 3).setValues([[st, new Date(), daftarSoal_().length]]);
  } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
