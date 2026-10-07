# Apps Script (backend)

Salinan kode Google Apps Script yang terikat ke spreadsheet budget. Folder ini **hanya arsip dan riwayat versi**: mengubah file di sini tidak mengubah apa pun yang berjalan.

| File | Isi |
|---|---|
| `Jembatan.gs` | Web App (`doGet`/`doPost`) yang dipanggil `index.html` |
| `SinkronSaham.gs` | Sinkron baris Beli/Jual di sheet Saham ke sheet Transaksi (pemicu `onEditSinkronSaham`) |
| `SnapshotSaham.gs` | Snapshot nilai saham ke sheet Aset (manual + pemicu harian `snapshotOtomatis`) |

Ketiga file harus ada di **satu proyek Apps Script yang sama**, karena saling memanggil (`trySyncRow_`, `tulisSnapshot_`).

## Menerapkan perubahan

1. Salin isi file yang berubah ke editor Apps Script (Extensions > Apps Script), lalu simpan.
2. Kalau `Jembatan.gs` berubah: Deploy > Manage deployments > edit deployment yang ada > Version: **New version** > Deploy. URL Web App tetap sama.

Token **tidak** disimpan di kode. Token ada di Project Settings > Script properties (`TOKEN`). Jangan pernah commit token atau URL Web App ke repo ini.
