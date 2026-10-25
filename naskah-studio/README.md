# Naskah Studio — Starter Apps Script

Fondasi aplikasi pengelolaan jasa edit dokumen berbasis Google Apps Script, Google Sheets, dan Google Drive. Tampilan web responsif memakai palet warna dari color guide.

## Isi

- `Code.gs` — pembuatan sheet, alur data, role, Drive, tugas, dan transaksi.
- `Index.html` — dashboard dan antarmuka pengguna.
- `appsscript.json` — manifest runtime V8 dan konfigurasi web app.

## Cara memasang

1. Buat Google Sheet baru untuk database.
2. Buka **Extensions → Apps Script**.
3. Buat file `Code.gs`, salin isi `Code.gs` dari paket ini. Tambahkan file HTML bernama `Index`, lalu salin `Index.html`.
4. Pada **Project Settings**, aktifkan tampilan `appsscript.json` dan salin manifest jika dibutuhkan. Atur zona waktu ke `Asia/Jakarta`.
5. Pilih fungsi `setup` lalu jalankan. Setujui akses Sheets dan Drive. Akun yang menjalankan setup pertama akan menjadi Owner.
6. Untuk pemakaian tim, deployment yang disarankan adalah **Execute as: Me** dengan akses dibatasi ke organisasi Google Workspace yang sama. Ini menjaga database dan Drive tetap berada di bawah akses owner, sementara aplikasi memeriksa email akun anggota tim. Jangan pilih akses publik.
7. Owner buka menu **Pengguna** untuk menambahkan email Google tim. Editor hanya akan melihat pesanan yang ditugaskan kepadanya.
8. Tambahkan katalog awal di menu **Layanan**. Daftar Excel layanan dapat dimasukkan setelah file diterima.

## Sheet yang dibuat

`Users`, `Clients`, `Services`, `Orders`, `OrderItems`, `Tasks`, `Payments`, dan `ActivityLog`. Fungsi `setup()` juga membuat folder Drive induk dan mencatat ID-nya pada Script Properties. Setiap pesanan memiliki folder dengan subfolder file awal, kerja, revisi, dan final.

## Batasan versi awal

- Unggah dari antarmuka dibatasi 8 MB per file.
- Pembatasan role yang aman memerlukan Google Workspace dengan anggota tim di domain yang sama dan deployment yang bisa memperoleh email pengguna. Pada akun Gmail pribadi atau deployment lintas domain, `Session.getActiveUser()` dapat kosong. Jangan gunakan aplikasi ini untuk data klien sebelum mekanisme identitas sesuai konfigurasi akun diverifikasi.
- Pengaturan visibilitas web app perlu ditetapkan sesuai kebijakan akun Google. Untuk data klien, jangan pilih akses publik.
- Belum ada pengiriman otomatis ke klien, tanda terima pembayaran, ekspor laporan, atau editor layanan/pesanan yang sudah tersimpan.
- Struktur layanan masih placeholder sampai Excel layanan diberikan.

## Catatan versi awal

Kode ini adalah fondasi untuk ditinjau dan disesuaikan dengan alur bisnis. Sebelum digunakan dengan data klien, pasang pada spreadsheet khusus dan coba dengan beberapa akun serta pesanan dummy untuk memastikan konfigurasi akses Drive sesuai kebutuhan.

