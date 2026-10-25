# haloeditin

Aplikasi pengelolaan pesanan jasa edit dokumen berbasis Apps Script, Google Sheets, dan Drive.

## Fitur saat ini

- Dashboard, pesanan, tugas editor, pembayaran, layanan, dan manajemen pengguna.
- Katalog sumber **List Harga Layanan Jasa EDITIN, pembaruan 1 Juni 2024**: 74 layanan, kategori, unit, keterangan, dan hingga sembilan tarif reguler/express.
- Harga item pesanan divalidasi kembali dari database server dan disimpan sebagai snapshot saat pesanan dibuat.
- Logo EDITIN disematkan pada halaman login dan navigasi.
- Owner membuat atau mengubah role, mengaktifkan akun, serta mengatur kata sandi awal minimal 10 karakter.
- Login menggunakan email dan kata sandi yang dibuat Owner; sesi berlaku enam jam. Owner juga dapat masuk memakai akun Google terdaftar.
- Kata sandi disimpan sebagai hash bersalt, percobaan login dibatasi, dan data tampilan editor dibatasi ke tugas yang ditugaskan.

## Penerapan

1. Salin `Code.gs`, `Index.html`, dan `appsscript.json` ke project Apps Script.
2. Jalankan `setup()` sebagai owner untuk menambahkan kolom database yang diperlukan.
3. Jalankan `importServiceCatalog()` sebagai owner. Fungsi ini aman dijalankan ulang tanpa membuat duplikat ID impor.
4. Masuk sebagai owner dan tambahkan akun editor/admin dari menu **Pengguna**.
5. Deploy sebagai web app yang berjalan sebagai owner. Untuk login staf, deployment perlu dapat dibuka oleh mereka; bagikan link hanya kepada staf yang sudah diberi akun. Semua endpoint data memerlukan sesi valid dan pengecekan role.

## Sheet dan Drive

Sheet: `Users`, `Clients`, `Services`, `Orders`, `OrderItems`, `Tasks`, `Payments`, dan `ActivityLog`. Setiap pesanan mendapat folder Drive dengan subfolder file awal, kerja, revisi, dan final. Batas unggah dari antarmuka 8 MB per file.

## Catatan

- Harga katalog mencerminkan file bertanggal 1 Juni 2024; tinjau kembali sebelum dijadikan tarif aktif bisnis.
- Apps Script tidak menyediakan sistem kata sandi khusus yang setara dengan penyedia identitas profesional. Untuk data sensitif atau pemakaian skala besar, gunakan Google Workspace Identity atau Firebase Authentication.
- Jangan ubah akses deployment sebelum kode versi baru disimpan dan endpoint terlindungi sesi; deployment yang dapat dibuka staf akan membuat halaman login dapat dijangkau melalui URL.
- Belum ada notifikasi otomatis, tanda terima, laporan ekspor, atau fitur mengedit item pesanan setelah dibuat.

