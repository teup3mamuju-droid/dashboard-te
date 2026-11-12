# Dashboard Penjualan Tenaga Listrik (TUL III-09)

Dashboard web interaktif dari file Excel Laporan III-09 (Lampiran 1), di-host gratis di GitHub Pages.

## Struktur
```
build_data.py              Excel -> docs/data/data.json
Transaksi Energi/309/      File Excel sumber (TIDAK di-upload ke GitHub)
  UP/<tahun>/...           level UP3
  KABUPATEN/<tahun>/...    level Kabupaten (otomatis muncul di dashboard bila ada file)
  UPI/<tahun>/...          level UPI
docs/                      Website (root GitHub Pages)
  index.html, assets/, data/data.json
```

Format file yang dikenali:
- **UP**: `Laporan_III_09_{NORMAL|LPB}_LAPORAN1_YYYYMM*.xlsx`
- **UPI (ditampilkan sebagai ULP)**: `Laporan_III_09_TOTAL_LAPORAN1_YYYYMM_<kode>.xlsx` — satu file per ULP.
  Nama ULP dibaca otomatis dari isi file; bisa juga diatur manual di `unit_names.json`
  (contoh: `{"32910": "ULP MANAKARRA"}`). Script mengecek jumlah semua ULP = UP3 setiap bulan.
- **KABUPATEN**: `TUL_III_09_KAB_<NAMA>_YYYYMM_TOTAL_YYYYMMDD.xlsx`

Detail untuk UP: file `Laporan_III_09_NORMAL_LAPORAN1_YYYYMM*.xlsx` (pascabayar) dan
`Laporan_III_09_LPB_LAPORAN1_YYYYMM*.xlsx` (prabayar). **Total = Normal + LPB** (dihitung otomatis,
sudah dicek sama persis dengan file TOTAL). Jika ada beberapa file untuk bulan yang sama
(mis. berakhiran `_04_Nov_2026_22-33-36`), yang dipakai adalah yang **terbaru**.
Untuk KABUPATEN/UPI: satu subfolder per unit/kabupaten bila nama unit di sel B3 sama.

## Update bulanan — lewat tombol di website (disarankan)
1. Buka **https://teup3mamuju-droid.github.io/dashboard-te/?admin** sekali saja di browser Anda.
   Tombol **Perbarui data** akan muncul (dan tetap muncul di browser itu). Pengunjung biasa tidak melihatnya.
2. Simpan file Excel baru ke folder `Transaksi Energi/309/...` seperti biasa.
3. Klik **Perbarui data → Pilih folder…** → pilih folder `309`.
   Browser membaca & memvalidasi semua file (logika sama dengan `build_data.py`, file tidak di-upload).
4. Periksa ringkasan & peringatan; dashboard menampilkan **pratinjau**.
5. Klik **Publikasikan**. Pertama kali diminta **token GitHub**:
   github.com → Settings → Developer settings → *Fine-grained tokens* → Generate new token →
   Repository access: *Only select repositories* → `dashboard-te` → Permissions: *Contents: Read and write*.
   Token disimpan hanya di browser tersebut.
6. Website ter-update untuk semua orang dalam ±1–2 menit.

Untuk menyembunyikan tombol lagi: buka `...?admin=0`.

> Jika setelah itu Anda ingin memakai cara script di bawah, jalankan `git pull` dulu
> karena data terbaru sudah dikirim langsung ke GitHub oleh tombol tadi.

## Update bulanan — lewat script (alternatif)
1. Simpan file Excel baru ke `Transaksi Energi/309/UP/2026/` (atau folder level lain).
2. Jalankan:
   ```
   python build_data.py
   ```
   Pastikan muncul `Validasi: semua total cocok dengan baris JUMLAH.`
3. Upload perubahan:
   ```
   git add docs/data/data.json
   git commit -m "Data bulan <bulan>"
   git push
   ```
   Website otomatis ter-update dalam 1–2 menit.

## Lihat di komputer sendiri
```
python -m http.server 8000 -d docs
```
lalu buka http://localhost:8000

## Setup GitHub Pages (sekali saja)
1. `gh auth login` (atau buat repo lewat github.com).
2. `gh repo create dashboard-te --public --source . --push`
3. Di GitHub: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `/docs`** → Save.
4. Alamat website: `https://<username>.github.io/dashboard-te/`

> Catatan: repo & website publik — siapa pun bisa melihat data agregat di `data.json`.
> File Excel mentah tidak ikut di-upload (lihat `.gitignore`).

## Fitur
- Filter Level / Unit / Tahun / Bulan / Jenis (Total, Pascabayar, Prabayar); link URL menyimpan filter (bisa dibagikan).
- Kartu KPI (pelanggan, daya, kWh, pendapatan, Rp/kWh) + perubahan vs bulan lalu + kumulatif + sparkline.
- Tren bulanan (klik batang untuk pilih bulan), komposisi golongan (klik untuk menyaring semua grafik),
  porsi pascabayar vs prabayar, 10 tarif terbesar, tabel detail (urut, cari, buka/tutup golongan, unduh CSV).
- Mode terang/gelap, tampilan responsif untuk HP.
