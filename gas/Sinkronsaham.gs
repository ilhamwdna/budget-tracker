/** @OnlyCurrentDoc */

/**
 * Menyinkronkan otomatis baris "Beli"/"Jual" di sheet Saham ke sheet Transaksi,
 * sesuai pemetaan yang sudah disepakati (Beli -> Tabungan & Investasi dengan kategori = "Tujuan untuk saham" di sheet Tujuan J2,
 * Jual -> Pemasukan/Hasil Investasi). Baris "Saldo awal" tidak disinkronkan.
 *
 * - siapkanSinkronSaham()     : jalankan SEKALI. Menambah kolom M "Rekening asal" (dropdown)
 *                               dan kolom O "Disalin ke Transaksi" di sheet Saham. Aman
 *                               dijalankan ulang (tidak menimpa kalau sudah ada).
 * - pasangTriggerSinkronSaham(): memasang pemicu otomatis: begitu kamu mengisi/mengubah
 *                               baris di Saham (dan kolom M sudah terisi), baris itu langsung
 *                               ditulis ke Transaksi.
 * - hapusTriggerSinkronSaham() : mencabut pemicu itu.
 * - sinkronSahamKeTransaksi()  : jalankan manual untuk menyisir SEMUA baris yang belum
 *                               tersinkron (misalnya baris lama sebelum pemicu dipasang).
 *
 * Catatan: sinkron hanya terjadi SEKALI per baris. Kolom O bertanda TRUE setelah tersalin;
 * mengedit baris itu lagi setelahnya TIDAK memperbarui baris Transaksi yang sudah tertulis.
 * Untuk koreksi, ubah baris Transaksi-nya secara langsung.
 */

const KOL_REKENING = 13; // M
const KOL_DISALIN = 15;  // O (L=modal keluar=12, N=catatan=14)

function siapkanSinkronSaham() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName('Saham');
  const master = SpreadsheetApp.getActive().getSheetByName('Master');
  if (!sh || !master) { ui.alert('Sheet Saham atau Master tidak ditemukan.'); return; }

  if (sh.getRange(1, KOL_REKENING).getValue() !== 'Rekening asal') {
    sh.getRange(1, KOL_REKENING).setValue('Rekening asal').setFontWeight('bold');
    sh.getRange(2, KOL_REKENING, 499, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(master.getRange('D2:D10'), true).setAllowInvalid(false).build()
    );
    sh.setColumnWidth(KOL_REKENING, 120);
  }
  if (sh.getRange(1, KOL_DISALIN).getValue() !== 'Disalin ke Transaksi') {
    sh.getRange(1, KOL_DISALIN).setValue('Disalin ke Transaksi').setFontWeight('bold');
    sh.setColumnWidth(KOL_DISALIN, 130);
  }
  ui.alert('Selesai. Kolom "Rekening asal" (M) dan "Disalin ke Transaksi" (O) sudah siap di Saham. Isi Rekening asal untuk baris Beli/Jual supaya bisa tersinkron.');
}

function pasangTriggerSinkronSaham() {
  hapusTriggerSinkronSaham();
  ScriptApp.newTrigger('onEditSinkronSaham').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getUi().alert('Pemicu dipasang. Setelah baris Beli/Jual di Saham lengkap (termasuk Rekening asal), baris itu otomatis tertulis ke Transaksi.');
}

function hapusTriggerSinkronSaham() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'onEditSinkronSaham')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

// Pemicu instalabel: dipanggil tiap ada edit di spreadsheet mana pun.
function onEditSinkronSaham(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Saham') return;
  const r0 = e.range.getRow(), r1 = e.range.getLastRow();
  const transaksi = SpreadsheetApp.getActive().getSheetByName('Transaksi');
  if (!transaksi) return;
  for (let r = Math.max(2, r0); r <= r1; r++) {
    trySyncRow_(sh, transaksi, r);
  }
}

// Menyisir semua baris (untuk baris lama / catch-up).
function sinkronSahamKeTransaksi() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Saham');
  const transaksi = ss.getSheetByName('Transaksi');
  if (!sh || !transaksi) { SpreadsheetApp.getUi().alert('Sheet Saham atau Transaksi tidak ditemukan.'); return; }
  const terakhir = sh.getLastRow();
  let tersinkron = 0, dilewati = 0;
  for (let r = 2; r <= terakhir; r++) {
    const hasil = trySyncRow_(sh, transaksi, r);
    if (hasil === 'ok') tersinkron++;
    else if (hasil === 'skip-incomplete') dilewati++;
  }
  SpreadsheetApp.getUi().alert(`Selesai. Baris baru tersinkron ke Transaksi: ${tersinkron}. Dilewati (belum lengkap / bukan Beli-Jual / sudah tersinkron): ${dilewati}.`);
}

// Mengecek satu baris Saham; menulis ke Transaksi bila memenuhi syarat dan belum tersinkron.
// Return: 'ok' | 'skip-incomplete' | 'skip-other'
function trySyncRow_(sahamSheet, transaksiSheet, row) {
  const v = sahamSheet.getRange(row, 1, 1, KOL_DISALIN).getValues()[0];
  const tanggal = v[0], kode = v[1], jenis = v[2], lembar = v[3], harga = v[4], biaya = v[5] || 0;
  const rekening = v[KOL_REKENING - 1];
  const sudahDisalin = v[KOL_DISALIN - 1] === true;

  if (sudahDisalin) return 'skip-other';
  if (jenis !== 'Beli' && jenis !== 'Jual') return 'skip-other'; // termasuk "Saldo awal" & baris kosong
  if (!(tanggal instanceof Date) || !kode || !(lembar > 0) || !(harga > 0) || !rekening) return 'skip-incomplete';

  let kategoriSaham = '';
  if (jenis === 'Beli') {
    // kategori setoran = tujuan untuk saham (Tujuan!J2); bila belum dipilih, baris belum disalin
    const ss = SpreadsheetApp.getActive();
    const tj = ss.getSheetByName('Tujuan'), master = ss.getSheetByName('Master');
    kategoriSaham = tj ? String(tj.getRange('J2').getValue()).trim() : '';
    // harus juga menjadi kategori Tabungan & Investasi di Master, kalau tidak Ringkasan tidak akan menghitungnya
    const valid = kategoriSaham && master && master.getRange('A2:B19').getValues().some((r) => String(r[0]).trim() === kategoriSaham && String(r[1]).trim() === 'Tabungan & Investasi');
    if (!valid) return 'skip-incomplete';
  }

  let jenisT, kategoriT, nilai;
  if (jenis === 'Beli') {
    jenisT = 'Tabungan & Investasi'; kategoriT = kategoriSaham; nilai = lembar * harga + biaya;
  } else {
    jenisT = 'Pemasukan'; kategoriT = 'Hasil Investasi'; nilai = lembar * harga - biaya;
  }
  const deskripsi = jenis + ' ' + kode;

  transaksiSheet.appendRow([tanggal, jenisT, kategoriT, deskripsi, nilai, rekening]);
  const barisBaru = transaksiSheet.getLastRow();
  transaksiSheet.getRange(barisBaru, 1).setNumberFormat('yyyy-mm-dd');
  transaksiSheet.getRange(barisBaru, 5).setNumberFormat('#,##0');

  sahamSheet.getRange(row, KOL_DISALIN).setValue(true);
  return 'ok';
}
