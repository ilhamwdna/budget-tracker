/** @OnlyCurrentDoc */

/**
 * Menyalin nilai pasar saham dari sheet Portofolio ke sheet Aset sebagai snapshot (angka beku),
 * satu baris per emiten dengan Jenis "Saham". Kolom Tujuan diisi otomatis dengan "tujuan untuk saham"
 * (sheet Tujuan, sel J2) yang berlaku saat snapshot ditulis, supaya baris ini langsung terhitung ke
 * tujuan itu di sheet Kekayaan tanpa perlu ditandai manual satu per satu. Bila J2 belum diisi (atau
 * berisi nama yang tidak ada di Master), kolom Tujuan dikosongkan seperti sebelumnya.
 *
 * - catatSnapshotSaham()  : jalankan manual kapan saja (menanyakan tanggal snapshot; menimpa baris
 *                           yang sudah ada untuk tanggal + emiten yang sama).
 * - pasangPemicuSnapshot(): pasang pemicu harian (sekitar jam 01.00). Pemicu hanya bertindak pada
 *                           tanggal 1-3 tiap bulan dan mencatat posisi akhir bulan sebelumnya
 *                           (termasuk posisi awal: akhir bulan sebelum periode dimulai),
 *                           hanya untuk baris yang belum ada (tidak menimpa).
 * - hapusPemicuSnapshot() : mencabut pemicu tersebut.
 *
 * Tidak memakai nama konstanta global agar tidak bentrok dengan file Saham.gs.
 */

function catatSnapshotSaham() {
  const ui = SpreadsheetApp.getUi();
  const jawab = ui.prompt('Tanggal snapshot', 'Format yyyy-mm-dd. Kosongkan untuk hari ini.', ui.ButtonSet.OK_CANCEL);
  if (jawab.getSelectedButton() !== ui.Button.OK) return;
  const teks = jawab.getResponseText().trim();
  let tgl = new Date();
  if (teks) {
    const m = teks.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) { ui.alert('Format tanggal salah. Contoh: 2026-10-31'); return; }
    tgl = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  tgl = new Date(tgl.getFullYear(), tgl.getMonth(), tgl.getDate(), 12, 0, 0); // jam 12 agar aman dari selisih zona waktu
  ui.alert(tulisSnapshot_(tgl, true));
}

// Dipanggil pemicu harian: hanya bertindak pada tanggal 1-3, untuk akhir bulan sebelumnya.
function snapshotOtomatis() {
  const sekarang = new Date();
  if (sekarang.getDate() > 3) return;
  const tgl = new Date(sekarang.getFullYear(), sekarang.getMonth(), 0, 12, 0, 0); // hari terakhir bulan lalu
  const master = SpreadsheetApp.getActive().getSheetByName('Master');
  const awalPeriode = master ? master.getRange('P2').getValue() : null;
  if (awalPeriode instanceof Date) {
    // Boleh mencatat posisi awal (akhir bulan sebelum periode dimulai), tapi tidak yang lebih awal dari itu.
    const batas = new Date(awalPeriode.getFullYear(), awalPeriode.getMonth() - 1, 1);
    if (tgl < batas) return;
  }
  console.log(tulisSnapshot_(tgl, false));
}

function pasangPemicuSnapshot() {
  hapusPemicuSnapshot();
  ScriptApp.newTrigger('snapshotOtomatis').timeBased().everyDays(1).atHour(1).create();
  SpreadsheetApp.getUi().alert('Pemicu dipasang: berjalan tiap hari sekitar jam 01.00 dan hanya bertindak pada tanggal 1-3 tiap bulan. Pastikan zona waktu proyek adalah Asia/Jakarta (Project Settings).');
}

function hapusPemicuSnapshot() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'snapshotOtomatis')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

// timpa=true: perbarui baris yang sudah ada. timpa=false: hanya tambah yang belum ada.
function tulisSnapshot_(tgl, timpa) {
  const ss = SpreadsheetApp.getActive();
  const porto = ss.getSheetByName('Portofolio');
  const aset = ss.getSheetByName('Aset');
  const tujSheet = ss.getSheetByName('Tujuan');
  const master = ss.getSheetByName('Master');
  if (!porto || !aset) return 'Sheet Portofolio atau Aset tidak ditemukan.';
  let tujuanSaham = '';
  if (tujSheet && master) {
    const calon = String(tujSheet.getRange('J2').getValue()).trim();
    const daftar = master.getRange('I2:I8').getValues().map((r) => String(r[0]).trim());
    if (calon && daftar.indexOf(calon) >= 0) tujuanSaham = calon;
  }

  SpreadsheetApp.flush();
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const kunci = (d) => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const target = kunci(tgl);

  const portoBaris = porto.getRange('A2:G9').getValues();           // A=kode, C=lembar, G=nilai pasar
  const terakhir = aset.getLastRow();
  const ada = terakhir >= 2 ? aset.getRange(2, 1, terakhir - 1, 5).getValues() : [];

  const baru = [], diperbarui = [], sudahAda = [], dilewati = [];
  portoBaris.forEach((r) => {
    const kode = String(r[0]).trim();
    if (!kode) return;
    if (!(Number(r[2]) > 0)) return;                                 // tidak dimiliki
    if (typeof r[6] !== 'number') { dilewati.push(kode); return; }   // harga belum tersedia
    const nama = 'Saham ' + kode;
    const idx = ada.findIndex((x) => x[0] instanceof Date && kunci(x[0]) === target && x[1] === nama);
    if (idx >= 0) {
      if (timpa) { aset.getRange(idx + 2, 4).setValue(tujuanSaham); aset.getRange(idx + 2, 5).setValue(r[6]); diperbarui.push(kode); }
      else sudahAda.push(kode);
      return;
    }
    baru.push([tgl, nama, 'Saham', tujuanSaham, r[6]]);
  });

  if (baru.length) {
    const rentang = aset.getRange(terakhir + 1, 1, baru.length, 5);
    rentang.setValues(baru);
    aset.getRange(terakhir + 1, 1, baru.length, 1).setNumberFormat('yyyy-mm-dd');
    aset.getRange(terakhir + 1, 5, baru.length, 1).setNumberFormat('#,##0');
  }
  let pesan = `Snapshot ${target}: ditambah ${baru.length}, diperbarui ${diperbarui.length}, sudah ada ${sudahAda.length}.`;
  if (dilewati.length) pesan += ` Dilewati karena harga belum tersedia: ${dilewati.join(', ')}.`;
  pesan += tujuanSaham ? ` Tujuan: ${tujuanSaham}.` : ' Tujuan belum diisi (Tujuan!J2 kosong atau tidak dikenal) — baris ditulis tanpa Tujuan.';
  return pesan;
}
