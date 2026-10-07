/** @OnlyCurrentDoc */

/**
 * JEMBATAN (Web App) — Tahap 3 bagian 1.
 * Menghubungkan aplikasi HTML dengan sheet. Semua nama diawali JB_ agar tidak bentrok
 * dengan file lain di proyek (Saham.gs, SinkronSaham.gs, dst).
 *
 * Aksi:
 *   GET  ?action=init&token=...              -> daftar kategori, rekening, emiten, bulan, dll (dari Master/Portofolio)
 *   GET  ?action=sisa&bulan=1..12&token=...  -> sisa budget pengeluaran per kategori untuk satu bulan
 *   POST {action:'catat', token, ...}        -> tulis satu transaksi ke sheet Transaksi
 *   POST {action:'catatSaham', token, ...}   -> tulis ke sheet Saham, lalu langsung disalin ke Transaksi
 *
 * Tambahan (versi 2), untuk layar PC:
 *   GET  ?action=transaksi&bulan=n           -> daftar transaksi satu bulan
 *   GET  ?action=budget&bulan=n              -> tabel budget (default + bulan terpilih + penanda diubah manual)
 *   GET  ?action=aset                        -> daftar snapshot aset
 *   GET  ?action=kekayaan                    -> tabel kekayaan (dibaca dari sheet Kekayaan)
 *   GET  ?action=saham                       -> tabel portofolio (dibaca dari sheet Portofolio)
 *   POST {action:'budgetSet', mode:'default'|'bulan'|'reset', jenis, kategori, bulan, nilai}
 *   POST {action:'asetTambah', tanggal, nama, jenis, tujuan, nilai}   (satu baris per akun per bulan: bila ada, diperbarui)
 *   POST {action:'masterTambah', tipe:'rekening'|'tujuan'|'emiten', nama, ket}
 *   POST {action:'kategoriSembunyi', nama, sembunyi:true|false}       (memakai Master kolom C)
 *   GET  ?action=tujuan                      -> progres tiap tujuan (dibaca dari sheet Tujuan)
 *   POST {action:'tujuanSet', nama, target, tanggal}                  (mengisi Target dan Tanggal target di sheet Tujuan)
 *   POST {action:'tujuanSet', tujuanSaham}                            (memilih tujuan untuk saham, sel J2)
 *   POST {action:'snapshotSaham', tanggal?, timpa?}                    (snapshot saham manual, lihat SnapshotSaham.gs)
 *
 * Token disimpan di Project Settings > Script properties (nama: TOKEN), bukan di kode ini.
 * Aksi menulis hanya diterima lewat POST.
 *
 * Catatan penting: pemicu onEdit TIDAK aktif untuk tulisan dari skrip/Web App (aturan Google),
 * jadi catatSaham memanggil trySyncRow_ (dari SinkronSaham.gs) secara langsung.
 */

// Zona waktu DIAMBIL DARI PENGATURAN SPREADSHEET SENDIRI (bukan di-hardcode), supaya tanggal yang
// ditulis skrip selalu sama dengan cara sheet menampilkannya, apa pun zona waktu yang dipakai sheet
// ini (File > Settings > Zona waktu). Sebelumnya di-hardcode 'Asia/Jakarta', dan kalau zona waktu
// sheet ternyata bukan itu, tanggal yang ditulis bisa bergeser mundur satu hari.
function JB_TZ_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
const JB_JENIS = ['Pengeluaran', 'Pemasukan', 'Tabungan & Investasi'];

function doGet(e) { return JB_handle_((e && e.parameter) || {}, 'GET'); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return JB_out_({ ok: false, error: 'Isi permintaan bukan JSON.' }); }
  return JB_handle_(body || {}, 'POST');
}

function JB_handle_(req, method) {
  try {
    if (!JB_authOk_(req.token)) return JB_out_({ ok: false, error: 'Tidak diizinkan.' });
    const aksi = String(req.action || '');
    if (method === 'GET') {
      if (aksi === 'init') return JB_out_(JB_init_());
      if (aksi === 'sisa') return JB_out_(JB_sisa_(req.bulan));
      if (aksi === 'transaksi') return JB_out_(JB_transaksi_(req.bulan));
      if (aksi === 'budget') return JB_out_(JB_budget_(req.bulan));
      if (aksi === 'aset') return JB_out_(JB_aset_());
      if (aksi === 'kekayaan') return JB_out_(JB_kekayaan_());
      if (aksi === 'saham') return JB_out_(JB_saham_());
      if (aksi === 'tujuan') return JB_out_(JB_tujuan_());
      if (aksi === 'cekZona') return JB_out_(JB_cekZona_());
      return JB_out_({ ok: false, error: 'Aksi GET tidak dikenal: ' + aksi });
    }
    if (aksi === 'catat') return JB_out_(JB_withLock_(() => JB_catat_(req)));
    if (aksi === 'catatSaham') return JB_out_(JB_withLock_(() => JB_catatSaham_(req)));
    if (aksi === 'budgetSet') return JB_out_(JB_withLock_(() => JB_budgetSet_(req)));
    if (aksi === 'asetTambah') return JB_out_(JB_withLock_(() => JB_asetTambah_(req)));
    if (aksi === 'masterTambah') return JB_out_(JB_withLock_(() => JB_masterTambah_(req)));
    if (aksi === 'kategoriSembunyi') return JB_out_(JB_withLock_(() => JB_kategoriSembunyi_(req)));
    if (aksi === 'tujuanSet') return JB_out_(JB_withLock_(() => JB_tujuanSet_(req)));
    if (aksi === 'snapshotSaham') return JB_out_(JB_withLock_(() => JB_snapshotSaham_(req)));
    return JB_out_({ ok: false, error: 'Aksi POST tidak dikenal: ' + aksi });
  } catch (err) {
    return JB_out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function JB_authOk_(token) {
  const benar = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!benar) throw new Error('TOKEN belum diatur di Script properties.');
  return typeof token === 'string' && token.length > 0 && token === benar;
}

function JB_out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function JB_withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- utilitas ----------
function JB_tgl_(d) { return Utilities.formatDate(d, JB_TZ_(), 'yyyy-MM-dd'); }

function JB_parseTgl_(teks) {
  if (typeof teks !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(teks)) throw new Error('Tanggal harus berformat yyyy-mm-dd.');
  const d = Utilities.parseDate(teks, JB_TZ_(), 'yyyy-MM-dd');
  if (JB_tgl_(d) !== teks) throw new Error('Tanggal tidak valid: ' + teks);
  return d;
}

function JB_daftarKategori_(master) {
  const peta = {};
  master.getRange('A2:B19').getValues().forEach((r) => { if (r[0]) peta[String(r[0])] = String(r[1]); });
  return peta; // { namaKategori: jenis }
}

function JB_daftarRekening_(master) {
  return master.getRange('D2:E10').getValues().filter((r) => r[0]).map((r) => ({ nama: String(r[0]), peran: String(r[1] || '') }));
}

// Tujuan untuk saham (Tujuan!J2): kategori setoran untuk pembelian saham. Kosong bila belum dipilih.
function JB_tujuanSaham_(ss) {
  const sh = ss.getSheetByName('Tujuan');
  return sh ? String(sh.getRange('J2').getValue()).trim() : '';
}

function JB_daftarEmiten_(porto) {
  return porto.getRange('A2:A9').getValues().map((r) => String(r[0]).trim()).filter(String);
}

function JB_periode_(master) {
  return master.getRange('M2:N13').getValues().map((r, i) => ({ no: i + 1, label: String(r[0]), mulai: JB_tgl_(r[1]) }));
}

// ---------- baca ----------
function JB_init_() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Master');
  const porto = ss.getSheetByName('Portofolio');

  const kategori = {}, kategoriSemua = {};
  JB_JENIS.forEach((j) => { kategori[j] = []; kategoriSemua[j] = []; });
  master.getRange('A2:C19').getValues().forEach((r) => {
    const n = String(r[0]), j = String(r[1]);
    if (!n || !kategoriSemua[j]) return;
    const hidden = String(r[2]).trim() !== '';
    kategoriSemua[j].push({ nama: n, hidden });
    if (!hidden) kategori[j].push(n);   // pilihan saat mencatat: tanpa yang disembunyikan
  });

  const rekening = JB_daftarRekening_(master);
  const cari = (kata) => { const r = rekening.filter((x) => x.peran.toLowerCase().indexOf(kata) >= 0)[0]; return r ? r.nama : null; };
  const rekeningDefault = { 'Pengeluaran': cari('harian'), 'Pemasukan': cari('gaji'), 'Tabungan & Investasi': null };

  const bulan = JB_periode_(master);
  const hariIni = JB_tgl_(new Date());
  let bulanIni = null;
  bulan.forEach((b, i) => {
    const akhir = i < bulan.length - 1 ? bulan[i + 1].mulai : null;
    if (hariIni >= b.mulai && (akhir === null || hariIni < akhir)) bulanIni = b.no;
  });

  const daftar = (rentang) => master.getRange(rentang).getValues().map((r) => String(r[0]).trim()).filter(String);
  return {
    ok: true,
    kategori, kategoriSemua, rekening, rekeningDefault, tujuanSaham: JB_tujuanSaham_(ss),
    tujuan: daftar('I2:I8'),
    jenisAset: daftar('G2:G9'),
    emiten: JB_daftarEmiten_(porto),
    bulan, bulanIni,
  };
}

// Menentukan bulan ke-n dalam periode (1..12) beserta batas tanggalnya (awal termasuk, akhir tidak).
function JB_rentang_(bulanParam) {
  const master = SpreadsheetApp.getActive().getSheetByName('Master');
  const bulan = JB_periode_(master);
  let idx = parseInt(bulanParam, 10);
  if (!(idx >= 1 && idx <= 12)) {
    const hariIni = JB_tgl_(new Date());
    idx = 1;
    bulan.forEach((b, i) => {
      const akhir = i < bulan.length - 1 ? bulan[i + 1].mulai : null;
      if (hariIni >= b.mulai && (akhir === null || hariIni < akhir)) idx = b.no;
    });
  }
  const awal = bulan[idx - 1].mulai;
  let akhir;
  if (idx < 12) akhir = bulan[idx].mulai;
  else { const p = awal.split('-').map(Number); akhir = Utilities.formatDate(new Date(p[0], p[1], 1, 12), JB_TZ_(), 'yyyy-MM-dd'); }
  return { idx, awal, akhir, label: bulan[idx - 1].label };
}

function JB_sisa_(bulanParam) {
  const ss = SpreadsheetApp.getActive();
  const rg = JB_rentang_(bulanParam);
  const idx = rg.idx, awal = rg.awal, akhir = rg.akhir;

  // budget bulan ini (Budget!A2:O22: A=jenis, B=kategori, D..O = bulan 1..12)
  const budget = ss.getSheetByName('Budget').getRange(2, 1, 21, 15).getValues();
  const anggaran = {};
  budget.forEach((r) => { if (r[1]) anggaran[String(r[0]) + '|' + String(r[1])] = Number(r[2 + idx]) || 0; });

  // realisasi dari Transaksi
  const trx = ss.getSheetByName('Transaksi');
  const n = trx.getLastRow();
  const real = { 'Pengeluaran': {}, 'Pemasukan': {}, 'Tabungan & Investasi': {} };
  if (n >= 2) {
    trx.getRange(2, 1, n - 1, 6).getValues().forEach((r) => {
      if (!(r[0] instanceof Date) || !real[r[1]]) return;
      const t = JB_tgl_(r[0]);
      if (t >= awal && t < akhir) real[r[1]][r[2]] = (real[r[1]][r[2]] || 0) + (Number(r[4]) || 0);
    });
  }
  const jumlah = (o) => Object.keys(o).reduce((s, k) => s + o[k], 0);

  const kategori = [];
  let totBudget = 0, totReal = 0;
  budget.forEach((r) => {
    if (r[0] !== 'Pengeluaran' || !r[1]) return;
    const b = Number(r[2 + idx]) || 0;
    const rl = real['Pengeluaran'][String(r[1])] || 0;
    totBudget += b; totReal += rl;
    kategori.push({ nama: String(r[1]), budget: b, realisasi: rl, sisa: b - rl });
  });

  const masuk = jumlah(real['Pemasukan']), keluar = jumlah(real['Pengeluaran']), tabung = jumlah(real['Tabungan & Investasi']);
  return {
    ok: true, bulan: idx, label: rg.label, awal, akhir,
    kategori,
    total: { budget: totBudget, realisasi: totReal, sisa: totBudget - totReal },
    arusKas: { pemasukan: masuk, pengeluaran: keluar, tabungan: tabung, belumTeralokasi: masuk - keluar - tabung },
  };
}

// ---------- tulis ----------
function JB_catat_(req) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Master');
  const tanggal = JB_parseTgl_(req.tanggal);
  const jenis = String(req.jenis || '');
  const kategori = String(req.kategori || '');
  const rekening = String(req.rekening || '');
  const nilai = Math.round(Number(req.nilai));
  const deskripsi = String(req.deskripsi || '').trim().slice(0, 200);

  if (JB_JENIS.indexOf(jenis) < 0) throw new Error('Jenis tidak dikenal.');
  const peta = JB_daftarKategori_(master);
  if (peta[kategori] !== jenis) throw new Error('Kategori "' + kategori + '" tidak cocok dengan jenis ' + jenis + '.');
  if (!JB_daftarRekening_(master).some((r) => r.nama === rekening)) throw new Error('Rekening tidak dikenal: ' + rekening);
  if (!(nilai > 0) || !isFinite(nilai)) throw new Error('Nilai harus angka lebih dari 0.');

  const trx = ss.getSheetByName('Transaksi');
  trx.appendRow([tanggal, jenis, kategori, deskripsi, nilai, rekening]);
  const baris = trx.getLastRow();
  trx.getRange(baris, 1).setNumberFormat('yyyy-mm-dd');
  trx.getRange(baris, 5).setNumberFormat('#,##0');
  return { ok: true, barisTransaksi: baris };
}

function JB_catatSaham_(req) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Master');
  const porto = ss.getSheetByName('Portofolio');
  const saham = ss.getSheetByName('Saham');
  const trx = ss.getSheetByName('Transaksi');

  const tanggal = JB_parseTgl_(req.tanggal);
  const kode = String(req.kode || '').trim().toUpperCase();
  const jenis = String(req.jenis || '');
  const lembar = Number(req.lembar);
  const harga = Number(req.harga);
  const biaya = req.biaya === undefined || req.biaya === '' ? 0 : Number(req.biaya);
  const rekening = String(req.rekening || '');

  if (jenis !== 'Beli' && jenis !== 'Jual') throw new Error('Jenis saham harus Beli atau Jual.');
  if (JB_daftarEmiten_(porto).indexOf(kode) < 0) throw new Error('Kode saham belum terdaftar di Portofolio: ' + kode);
  if (!(lembar > 0) || Math.floor(lembar) !== lembar) throw new Error('Lembar harus bilangan bulat lebih dari 0.');
  if (!(harga > 0) || !isFinite(harga)) throw new Error('Harga per lembar harus lebih dari 0.');
  if (!(biaya >= 0) || !isFinite(biaya)) throw new Error('Biaya tidak valid.');
  if (!JB_daftarRekening_(master).some((r) => r.nama === rekening)) throw new Error('Rekening tidak dikenal: ' + rekening);

  if (jenis === 'Beli') {
    const katSaham = JB_tujuanSaham_(ss);
    if (!katSaham) throw new Error('Tujuan untuk saham belum dipilih. Pilih di Setelan > Tujuan aset (atau sheet Tujuan, sel J2).');
    if (JB_daftarKategori_(master)[katSaham] !== 'Tabungan & Investasi') throw new Error('Tujuan untuk saham ("' + katSaham + '") bukan kategori Tabungan & Investasi.');
  }
  if (jenis === 'Jual') {
    SpreadsheetApp.flush();
    const dimiliki = porto.getRange('A2:C9').getValues().filter((r) => String(r[0]).trim() === kode).reduce((s, r) => s + (Number(r[2]) || 0), 0);
    if (lembar > dimiliki) throw new Error('Lembar yang dijual (' + lembar + ') melebihi kepemilikan (' + dimiliki + ').');
  }

  // baris kosong pertama (kolom Kode) — rumus G:L sudah terisi sampai baris 500, jadi jangan pakai appendRow
  const kolomKode = saham.getRange(2, 2, 499, 1).getValues();
  let baris = -1;
  for (let i = 0; i < kolomKode.length; i++) { if (String(kolomKode[i][0]).trim() === '') { baris = i + 2; break; } }
  if (baris < 0) throw new Error('Sheet Saham penuh (500 baris).');

  saham.getRange(baris, 1, 1, 6).setValues([[tanggal, kode, jenis, lembar, harga, biaya || '']]);
  saham.getRange(baris, 1).setNumberFormat('yyyy-mm-dd');
  saham.getRange(baris, 13).setValue(rekening);   // M = Rekening asal
  saham.getRange(baris, 15).setValue('');         // O = tanda "Disalin"; bersihkan sisa lama sebelum sinkron
  SpreadsheetApp.flush();

  const hasil = trySyncRow_(saham, trx, baris);
  if (hasil !== 'ok') {
    throw new Error('Tersimpan di Saham baris ' + baris + ' tetapi gagal disalin ke Transaksi (status: ' + hasil + '). Cek baris itu di sheet.');
  }
  const nilaiTransaksi = jenis === 'Beli' ? lembar * harga + biaya : lembar * harga - biaya;
  return { ok: true, barisSaham: baris, barisTransaksi: trx.getLastRow(), nilaiTransaksi };
}


// ---------- layar PC: baca ----------
function JB_transaksi_(bulanParam) {
  const rg = JB_rentang_(bulanParam);
  const trx = SpreadsheetApp.getActive().getSheetByName('Transaksi');
  const n = trx.getLastRow();
  const baris = [];
  if (n >= 2) {
    trx.getRange(2, 1, n - 1, 6).getValues().forEach((r, i) => {
      if (!(r[0] instanceof Date)) return;
      const t = JB_tgl_(r[0]);
      if (t >= rg.awal && t < rg.akhir) {
        baris.push({ baris: i + 2, tanggal: t, jenis: String(r[1]), kategori: String(r[2]), deskripsi: String(r[3] || ''), nilai: Number(r[4]) || 0, rekening: String(r[5] || '') });
      }
    });
  }
  baris.sort((a, b) => (a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : b.baris - a.baris));
  return { ok: true, bulan: rg.idx, label: rg.label, baris };
}

function JB_budget_(bulanParam) {
  const rg = JB_rentang_(bulanParam);
  const sh = SpreadsheetApp.getActive().getSheetByName('Budget');
  const nilai = sh.getRange(2, 1, 21, 15).getValues();
  const rumus = sh.getRange(2, 4, 21, 12).getFormulas();   // D..O: kosong = angka diketik manual (ditimpa)
  const baris = [];
  nilai.forEach((r, i) => {
    if (!r[1]) return;
    const isSub = String(r[0]) === '';
    baris.push({
      row: i + 2, jenis: String(r[0]), kategori: String(r[1]), isSub,
      default: Number(r[2]) || 0, nilai: Number(r[2 + rg.idx]) || 0,
      override: !isSub && String(rumus[i][rg.idx - 1]) === '',
    });
  });
  return { ok: true, bulan: rg.idx, label: rg.label, baris };
}

function JB_aset_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Aset');
  const n = sh.getLastRow();
  const baris = [];
  if (n >= 2) {
    sh.getRange(2, 1, n - 1, 5).getValues().forEach((r, i) => {
      if (!(r[0] instanceof Date)) return;
      baris.push({ baris: i + 2, tanggal: JB_tgl_(r[0]), nama: String(r[1]), jenis: String(r[2]), tujuan: String(r[3] || ''), nilai: Number(r[4]) || 0 });
    });
  }
  baris.sort((a, b) => (a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : b.baris - a.baris));
  return { ok: true, baris: baris.slice(0, 300) };
}

function JB_kekayaan_() {
  const v = SpreadsheetApp.getActive().getSheetByName('Kekayaan').getRange(4, 1, 14, 18).getValues();
  const ubah = (x) => (x instanceof Date ? JB_tgl_(x) : x);
  // hanya kolom yang punya judul, supaya kolom yang sudah dihapus dari sheet tidak muncul sebagai kolom kosong
  const pakai = [];
  v[0].forEach((h, i) => { if (String(h).trim() !== '') pakai.push(i); });
  return {
    ok: true,
    header: pakai.map((i) => String(v[0][i])),
    baris: v.slice(1).map((r) => pakai.map((i) => ubah(r[i]))),
  };
}

function JB_saham_() {
  const v = SpreadsheetApp.getActive().getSheetByName('Portofolio').getRange(2, 1, 9, 11).getValues();  // baris 2..10
  const angka = (x) => (typeof x === 'number' ? x : (x === '' ? null : String(x)));
  const baris = [];
  v.slice(0, 8).forEach((r) => {
    if (String(r[0]).trim() === '') return;
    baris.push({ kode: String(r[0]).trim(), lembar: angka(r[2]), rata: angka(r[3]), modal: angka(r[4]), sekarang: angka(r[5]),
                 pasar: angka(r[6]), ur: angka(r[7]), urPct: angka(r[8]), bobot: angka(r[9]), real: angka(r[10]) });
  });
  const t = v[8];
  return { ok: true, baris, total: { modal: angka(t[4]), pasar: angka(t[6]), ur: angka(t[7]), urPct: angka(t[8]), real: angka(t[10]) } };
}

// ---------- layar PC: tulis ----------
function JB_barisKosong_(sheet, kolom, dari, sampai) {
  const v = sheet.getRange(dari, kolom, sampai - dari + 1, 1).getValues();
  for (let i = 0; i < v.length; i++) { if (String(v[i][0]).trim() === '') return dari + i; }
  return -1;
}

function JB_budgetSet_(req) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Budget');
  const mode = String(req.mode || '');
  const bulan = parseInt(req.bulan, 10);
  const jenis = String(req.jenis || ''), kategori = String(req.kategori || '');
  if (['default', 'bulan', 'reset'].indexOf(mode) < 0) throw new Error('Mode budget tidak dikenal.');
  if (mode !== 'default' && !(bulan >= 1 && bulan <= 12)) throw new Error('Bulan harus 1 sampai 12.');
  const kunci = sh.getRange(2, 1, 21, 2).getValues();
  let row = -1;
  kunci.forEach((r, i) => { if (row < 0 && String(r[0]) === jenis && String(r[1]) === kategori && jenis !== '') row = i + 2; });
  if (row < 0) throw new Error('Kategori budget tidak ditemukan: ' + kategori);
  if (mode === 'reset') {
    sh.getRange(row, 3 + bulan).setFormula('=$C' + row);
    return { ok: true };
  }
  const nilai = Math.round(Number(req.nilai));
  if (!isFinite(nilai) || nilai < 0) throw new Error('Nilai budget harus angka 0 atau lebih.');
  if (mode === 'default') sh.getRange(row, 3).setValue(nilai);
  else sh.getRange(row, 3 + bulan).setValue(nilai);
  return { ok: true };
}

function JB_asetTambah_(req) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Master');
  const sh = ss.getSheetByName('Aset');
  const tanggal = JB_parseTgl_(req.tanggal);
  const nama = String(req.nama || '').trim().slice(0, 60);
  const jenis = String(req.jenis || '');
  let tujuan = String(req.tujuan || '').trim();
  const nilai = Math.round(Number(req.nilai));
  const daftar = (rentang) => master.getRange(rentang).getValues().map((r) => String(r[0]).trim()).filter(String);
  if (!nama) throw new Error('Nama akun/aset wajib diisi.');
  if (daftar('G2:G9').indexOf(jenis) < 0) throw new Error('Jenis aset tidak dikenal: ' + jenis);
  if (tujuan && daftar('I2:I8').indexOf(tujuan) < 0) throw new Error('Tujuan tidak dikenal: ' + tujuan);
  if (!isFinite(nilai) || nilai < 0) throw new Error('Nilai harus angka 0 atau lebih.');

  const bulanIni = JB_tgl_(tanggal).slice(0, 7);
  const v = sh.getRange(2, 1, 999, 2).getValues();
  let baris = -1, diperbarui = false;
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] instanceof Date && JB_tgl_(v[i][0]).slice(0, 7) === bulanIni && String(v[i][1]).trim().toLowerCase() === nama.toLowerCase()) { baris = i + 2; diperbarui = true; break; }
  }
  if (baris < 0) {
    baris = JB_barisKosong_(sh, 1, 2, 1000);
    if (baris < 0) throw new Error('Sheet Aset penuh.');
  }
  if (diperbarui && !tujuan) tujuan = String(sh.getRange(baris, 4).getValue() || '').trim();   // kosong = pertahankan tujuan lama
  sh.getRange(baris, 1, 1, 5).setValues([[tanggal, nama, jenis, tujuan, nilai]]);
  sh.getRange(baris, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(baris, 5).setNumberFormat('#,##0');
  return { ok: true, diperbarui, baris, tujuan };
}

function JB_masterTambah_(req) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Master');
  const tipe = String(req.tipe || '');
  let nama = String(req.nama || '').trim();
  const ket = String(req.ket || '').trim().slice(0, 60);
  if (!nama) throw new Error('Nama wajib diisi.');
  if (nama.length > 40) throw new Error('Nama terlalu panjang (maksimal 40 karakter).');

  let sheet, kolom, dari, sampai;
  if (tipe === 'rekening') { sheet = master; kolom = 4; dari = 2; sampai = 10; }
  else if (tipe === 'tujuan') { sheet = master; kolom = 9; dari = 2; sampai = 8; }
  else if (tipe === 'emiten') {
    sheet = ss.getSheetByName('Portofolio'); kolom = 1; dari = 2; sampai = 9;
    nama = nama.toUpperCase();
    if (!/^[A-Z]{2,6}$/.test(nama)) throw new Error('Kode emiten harus 2 sampai 6 huruf, tanpa awalan IDX.');
  } else throw new Error('Jenis pengaturan tidak dikenal: ' + tipe);

  const ada = sheet.getRange(dari, kolom, sampai - dari + 1, 1).getValues().some((r) => String(r[0]).trim().toLowerCase() === nama.toLowerCase());
  if (ada) throw new Error(nama + ' sudah ada.');
  const baris = JB_barisKosong_(sheet, kolom, dari, sampai);
  if (baris < 0) throw new Error('Slot penuh. Perlu diperluas dari sheet.');
  sheet.getRange(baris, kolom).setValue(nama);
  if (tipe === 'rekening') master.getRange(baris, 5).setValue(ket);
  return { ok: true, baris };
}

function JB_kategoriSembunyi_(req) {
  const master = SpreadsheetApp.getActive().getSheetByName('Master');
  const nama = String(req.nama || '');
  const sembunyi = req.sembunyi === true;
  const v = master.getRange('A2:C19').getValues();
  let idx = -1;
  v.forEach((r, i) => { if (idx < 0 && String(r[0]) === nama) idx = i; });
  if (idx < 0) throw new Error('Kategori tidak ditemukan: ' + nama);
  if (sembunyi) {
    const jenis = String(v[idx][1]);
    const sisa = v.filter((r, i) => i !== idx && String(r[1]) === jenis && String(r[2]).trim() === '').length;
    if (sisa < 1) throw new Error('Tidak bisa menyembunyikan kategori terakhir untuk jenis ' + jenis + '.');
  }
  if (String(master.getRange('C1').getValue()).trim() === '') master.getRange('C1').setValue('Sembunyikan');
  master.getRange(idx + 2, 3).setValue(sembunyi ? 'ya' : '');
  return { ok: true };
}

// ---------- tujuan ----------
function JB_tujuan_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Tujuan');
  if (!sh) throw new Error('Sheet "Tujuan" belum dibuat. Jalankan buatSheetTujuan di Apps Script.');
  const v = sh.getRange(2, 1, 9, 7).getValues();                 // baris 2..10
  const num = (x) => (typeof x === 'number' ? x : null);
  const baris = [];
  v.slice(0, 7).forEach((r) => {
    const nama = String(r[0]).trim();
    if (!nama) return;
    baris.push({
      nama, target: num(r[1]), tanggalTarget: r[2] instanceof Date ? JB_tgl_(r[2]) : null,
      nilai: num(r[3]) || 0, progres: num(r[4]), sisa: num(r[5]),
      perlu: typeof r[6] === 'number' ? r[6] : (r[6] === 'lewat' ? 'lewat' : null),
    });
  });
  const par = sh.getRange('J1:J5').getValues();
  // Awal/Akhir bulan acuan (J3/J4): dikirim supaya klien bisa membatasi data Aset mentah (endpoint 'aset', yang tidak
  // difilter per bulan) ke rentang yang SAMA dengan yang dipakai rumus di sheet ini, saat menghitung rincian jenis
  // aset per tujuan. Tanpa ini, rincian itu bisa menjumlahkan snapshot dari banyak bulan sekaligus.
  const awalAcuan = par[2][0] instanceof Date ? JB_tgl_(par[2][0]) : '';
  const akhirAcuan = par[3][0] instanceof Date ? JB_tgl_(par[3][0]) : '';
  return { ok: true, bulanAcuan: String(par[4][0]), awalAcuan, akhirAcuan, tujuanSaham: String(par[1][0]).trim(), baris, tanpaTujuan: num(v[7][3]) || 0, total: num(v[8][3]) || 0 };
}

function JB_tujuanSet_(req) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Tujuan');
  if (!sh) throw new Error('Sheet "Tujuan" belum dibuat. Jalankan buatSheetTujuan di Apps Script.');
  if (req.tujuanSaham !== undefined) {
    const nama = String(req.tujuanSaham).trim();
    const master = ss.getSheetByName('Master');
    if (nama !== '' && master.getRange('I2:I8').getValues().map((r) => String(r[0]).trim()).indexOf(nama) < 0) throw new Error('Tujuan tidak dikenal: ' + nama);
    if (nama !== '' && JB_daftarKategori_(master)[nama] !== 'Tabungan & Investasi') {
      throw new Error('Tujuan untuk saham harus juga menjadi kategori Tabungan & Investasi (pembelian saham dicatat ke kategori itu).');
    }
    sh.getRange('J2').setValue(nama);
    return { ok: true };
  }
  const nama = String(req.nama || '').trim();
  const v = sh.getRange(2, 1, 7, 1).getValues();
  let row = -1;
  v.forEach((r, i) => { if (row < 0 && nama !== '' && String(r[0]).trim() === nama) row = i + 2; });
  if (row < 0) throw new Error('Tujuan tidak ditemukan: ' + nama);
  let target = '';
  if (req.target !== '' && req.target != null) {
    target = Math.round(Number(req.target));
    if (!isFinite(target) || target < 0) throw new Error('Target harus angka 0 atau lebih.');
  }
  const tanggal = req.tanggal ? JB_parseTgl_(req.tanggal) : '';
  sh.getRange(row, 2).setValue(target);
  sh.getRange(row, 3).setValue(tanggal);
  sh.getRange(row, 3).setNumberFormat('yyyy-mm-dd');
  return { ok: true };
}

// ---------- snapshot saham manual (tombol di aplikasi) ----------
// Memanggil tulisSnapshot_ dari SnapshotSaham.gs secara LANGSUNG (fungsi di proyek Apps Script yang sama
// bisa saling panggil tanpa apa pun). File SnapshotSaham.gs harus ada di proyek yang sama dengan file ini.
function JB_snapshotSaham_(req) {
  const tz = JB_TZ_();
  let teks = String(req.tanggal || '').trim();
  if (!teks) teks = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(teks)) throw new Error('Tanggal harus berformat yyyy-mm-dd.');
  const tgl = Utilities.parseDate(teks + ' 12:00:00', tz, 'yyyy-MM-dd HH:mm:ss');   // jam 12 agar aman dari selisih zona waktu
  if (Utilities.formatDate(tgl, tz, 'yyyy-MM-dd') !== teks) throw new Error('Tanggal tidak valid: ' + teks);
  const timpa = req.timpa !== false;   // bawaan: tombol manual = "perbarui sekarang", boleh menimpa yang sudah ada
  if (typeof tulisSnapshot_ !== 'function') {
    throw new Error('Fungsi tulisSnapshot_ tidak ditemukan. Pastikan SnapshotSaham.gs sudah ditempel di proyek Apps Script yang sama dengan Jembatan.gs.');
  }
  return { ok: true, pesan: tulisSnapshot_(tgl, timpa) };
}

// Diagnostik, bisa dibuka langsung di browser: URL?action=cekZona&token=TOKENMU
// Menguji zona waktu PERSIS lewat jalur yang sama dengan yang dipakai aplikasi (doGet Web App),
// bukan lewat editor Apps Script, supaya ketahuan kalau keduanya ternyata berbeda.
function JB_cekZona_() {
  const sheetTz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const scriptTz = Session.getScriptTimeZone();
  const dipakai = JB_TZ_();
  const tes = '2026-10-01';
  const tgl = JB_parseTgl_(tes);
  const kembali = JB_tgl_(tgl);
  return {
    ok: true,
    zonaWaktuSheet: sheetTz || '(kosong)',
    zonaWaktuProyekScript: scriptTz || '(kosong)',
    zonaWaktuDipakai: dipakai,
    tesTanggal: tes,
    tesHasil: kembali,
    tesCocok: kembali === tes,
    tesDetailJam: Utilities.formatDate(tgl, dipakai, "yyyy-MM-dd HH:mm 'zona' z"),
  };
}

// ---------- uji manual dari editor Apps Script ----------
function JB_uji() {
  const a = JB_init_();
  console.log('init: kategori=' + JSON.stringify(a.kategori));
  console.log('rekening=' + JSON.stringify(a.rekening) + ' default=' + JSON.stringify(a.rekeningDefault));
  console.log('emiten=' + JSON.stringify(a.emiten) + ' bulanIni=' + a.bulanIni);
  const s = JB_sisa_(a.bulanIni || 1);
  console.log('sisa ' + s.label + ': ' + JSON.stringify(s.total) + ' arusKas=' + JSON.stringify(s.arusKas));
}

// Diagnostik sementara: jalankan dari editor Apps Script (bukan lewat aplikasi), lalu lihat hasilnya di
// View > Executions atau tab Log. Menunjukkan zona waktu yang sebenarnya dipakai Jembatan.gs sekarang.
function JB_cekZonaWaktu() {
  const sheetTz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const scriptTz = Session.getScriptTimeZone();
  const dipakai = JB_TZ_();
  const contoh = Utilities.formatDate(Utilities.parseDate('2026-10-01', dipakai, 'yyyy-MM-dd'), dipakai, "yyyy-MM-dd HH:mm 'zona' z");
  const pesan = 'Zona waktu sheet: ' + sheetTz + '\nZona waktu proyek script: ' + scriptTz +
    '\nZona waktu yang DIPAKAI Jembatan.gs sekarang: ' + dipakai +
    '\n\n"2026-10-01" ditulis ulang dan dibaca kembali sebagai: ' + contoh;
  Logger.log(pesan);
  SpreadsheetApp.getUi().alert(pesan);
}
