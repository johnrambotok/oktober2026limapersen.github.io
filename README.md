# IDX Ownership Monitor (5%+)

Dashboard statis untuk membaca laporan XLSX kepemilikan efek di atas 5% berdasarkan SID.

## Cara menggunakan
1. Buka repository → **Settings → Pages**.
2. Pada **Build and deployment → Source**, pilih **GitHub Actions**, lalu simpan.
3. Buka tab **Actions** dan jalankan ulang workflow **Deploy dashboard to GitHub Pages** yang gagal sebelumnya (Re-run jobs). Setelah itu, setiap commit ke `main` akan memicu deployment otomatis.
4. Pilih file laporan pembanding pada input pertama dan file laporan terbaru pada input kedua.
5. Klik **Analisis laporan**.
6. Gunakan pencarian dan filter untuk mengeksplorasi data; klik header tabel untuk mengurutkan; gunakan **Ekspor CSV** untuk menyimpan hasil filter.

## Catatan interpretasi
- File Excel diproses di browser pengguna; dashboard tidak mengunggah file ke server.
- Untuk membandingkan laporan, gunakan file yang tanggal laporannya berbeda.
- Baris yang tidak ditemukan di laporan terbaru ditandai sebagai tidak ditemukan; bukan otomatis dianggap kepemilikan nol.
- Data hanya mencakup laporan yang dipilih dan bukan rekomendasi investasi.

## File
- `index.html`: UI dashboard
- `app.js`: parser XLSX, filter, grafik, tabel, ekspor
- `style.css`: tampilan responsif
