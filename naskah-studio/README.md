# Naskah Studio — Starter Apps Script

Fondasi aplikasi pengelolaan jasa edit dokumen berbasis Google Apps Script, Google Sheets, dan Google Drive. Tampilan web responsif memakai palet warna dari color guide.

## Isi

- `Code.gs` — pembuatan sheet, alur data, role, Drive, tugas, dan transaksi.
- `Index.html` — dashboard dan antarmuka pengguna.
- `appsscript.json` — manifest runtime V8 dan konfigurasi web app.

## Cara memasang

1. Buka proyek Apps Script baru, buat file `Code.gs` dan HTML bernama `Index`, lalu salin file dari paket ini.
2. Jalankan `setup()`. Fungsi ini membuat Google Sheet database dan folder induk Drive; akun pertama menjadi Owner.
3. Untuk pemakaian tim, pilih skema deployment dan akses setelah metode identitas akun dipastikan. Jangan pilih akses publik.
4. Tambahkan email tim dari menu **Pengguna** setelah deployment yang aman tersedia. Editor hanya akan melihat tugas yang ditugaskan kepadanya.
5. Tambahkan atau impor katalog layanan setelah file Excel diterima.

## Sheet yang dibuat

`Users`, `Clients`, `Services`, `Orders`, `OrderItems`, `Tasks`, `Payments`, dan `ActivityLog`. Fungsi `setup()` juga membuat folder Drive induk dan mencatat ID-nya pada Script Properties. Setiap pesanan memiliki folder dengan subfolder file awal, kerja, revisi, dan final.

## Batasan versi awal

- Unggah dari antarmuka dibatasi 8 MB per file.
- Pembatasan role yang aman memerlukan deployment yang bisa memperoleh email pengguna. Pada akun Gmail pribadi atau deployment lintas domain, `Session.getActiveUser()` dapat kosong. Mode **Execute as: User accessing** dapat memberi setiap pengguna akses langsung ke spreadsheet dan berkas Drive; evaluasi konsekuensi itu sebelum dipakai bersama.
- Pengaturan visibilitas web app perlu ditetapkan sesuai kebijakan akun Google. Untuk data klien, jangan pilih akses publik.
- Belum ada pengiriman otomatis ke klien, tanda terima pembayaran, ekspor laporan, atau editor layanan/pesanan yang sudah tersimpan.
- Struktur layanan masih placeholder sampai Excel layanan diberikan.

## Catatan versi awal

Kode ini adalah fondasi untuk ditinjau dan disesuaikan dengan alur bisnis. Sebelum digunakan dengan data klien, pasang pada spreadsheet khusus dan coba dengan beberapa akun serta pesanan dummy untuk memastikan konfigurasi akses Drive sesuai kebutuhan.

