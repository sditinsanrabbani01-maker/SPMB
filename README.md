# SPMB Online - Sistem Pendaftaran Baru

## Deskripsi
SPMB Online adalah sistem pendaftaran peserta didik baru untuk SDIT Insan Rabbani. Sistem ini terdiri dari beberapa halaman statis yang terhubung dengan Google Apps Script sebagai backend untuk menyimpan data pendaftaran, memproses pembayaran, dan mengelola pengaturan.

## Struktur File
- `index.html` - Halaman beranda dengan informasi tentang sekolah, visi & misi, dan galeri kegiatan.
- `form.html` - Halaman formulir pendaftaran siswa baru.
- `pembayaran.html` - Halaman untuk upload bukti pembayaran dan verifikasi otomatis menggunakan Tesseract.js.
- `cek.html` - Halaman untuk cek status pendaftaran menggunakan nomor registrasi/NISN.
- `admin.html` - Dashboard admin untuk melihat data pendaftar, mengubah status, dan mengatur pengaturan sistem.
- `closed.html` - Halaman yang ditampilkan ketika pendaftaran ditutup.
- `login.html` - Halaman login untuk mengakses dashboard admin (sederhana, menggunakan hardcoded credentials).
- `config.js` - File konfigurasi yang berisi URL Google Apps Script dan konstanta lain.
- `kode.gs` - Google Apps Script yang berisi logika backend untuk menangani permintaan dari frontend.
- `*.jpg`, `*.jpeg` - Gambar yang digunakan di seluruh situs (banner, galeri, dll).
- `CNAME` - File untuk konfigurasi custom domain (jika digunakan).

## Cara Kerja
1. Pengunjung mengakses `index.html` dan dapat mendaftar melalui tombol "Daftar Sekarang" yang mengarahkan ke `pembayaran.html`.
2. Di `pembayaran.html`, pengguna mengupload bukti pembayaran. Sistem menggunakan Tesseract.js untuk membaca jumlah pembayaran dari gambar dan memverifikasi apakah minimal Rp 100.000.
3. Jika pembayaran valid, pengguna diarahkan ke `form.html` untuk mengisi formulir pendaftaran.
4. Setelah formulir disimpan, data dikirim ke Google Apps Script (`kode.gs`) untuk disimpan ke spreadsheet dan nomor registrasi dikembalikan.
5. Pengguna dapat cek status pendaftaran di `cek.html` menggunakan nomor registrasi atau NISN.
6. Admin dapat login melalui `login.html` (username: admin, password: admin123 - disarankan untuk diubah) dan mengakses `admin.html` untuk mengelola data pendaftar dan pengaturan sistem.

## Teknologi yang Digunakan
- HTML5, CSS3, Bootstrap 5
- Font Awesome untuk ikon
- Tesseract.js untuk OCR (pembacaan bukti pembayaran)
- Google Apps Script sebagai backend dan database (Google Sheets)
- Vanilla JavaScript untuk logika frontend

## Cara Menjalankan
Karena ini adalah situs statis dengan backend Google Apps Script, Anda hanya perlu:
1. Menyimpan semua file ke dalam sebuah folder yang dapat diakses oleh web server (misalnya, GitHub Pages, Netlify, atau hosting biasa).
2. Pastikan `config.js` berisi URL yang benar ke Google Apps Script deployment Anda.
3. Deploy `kode.gs` sebagai Google Apps Script dengan tipe deployment "Web app" dan set akses siapa saja, bahkan anonimitas.
4. Pastikan spreadsheet yang terkait memiliki sheet dengan nama "Data_SPMB", "Settings", dan "Pembayaran" (akan dibuat otomatis oleh script jika tidak ada).

## Catatan Keamanan
- Kredensial admin di `login.html` dan `admin.html` masih menggunakan hardcoded username dan password. Disarankan untuk mengganti dengan sistem autentikasi yang lebih aman atau setidaknya mengganti password di `config.js`.
- Tidak ada sanitasi input yang luas di frontend; namun, Google Apps Script harus memvalidasi data sebelum disimpan.

## Penambahan yang Disarankan
1. **Autentikasi yang lebih aman** - Ganti sistem login sederhana dengan autentikasi yang terenkripsi atau gunakan layanan seperti Firebase Authentication.
2. **Validasi Input yang Lebih Komprehensif** - Tambahkan validasi di frontend dan backend untuk mencegah data yang tidak valid.
3. **Notifikasi Email** - Selain WhatsApp, tambahkan notifikasi email pendaftaran berhasil.
4. **Halaman 404** - Buat halaman 404 kustom untuk menangani URL yang tidak ditemukan.
5. **SEO dan Meta Tags** - Tambahkan meta tag untuk deskripsi dan kata kunci agar lebih mudah ditemukan oleh mesin pencari.
6. **Penerjemahan Bahasa** - Situs saat ini hanya dalam Bahasa Indonesia; pertimbangkan untuk menambahkan dukungan bahasa Inggris jika diperlukan.
7. **Analytics** - Integrasi Google Analytics atau analisis serupa untuk melacak pengunjung.
8. **Peningkatan Aksesibilitas** - Pastikan semua elemen sesuai dengan standar WCAG (misalnya, label yang tepat, kontras warna yang cukup).
9. **Upload Bukti Pembayaran yang Lebih Robust** - Tambahkan preview gambar sebelum upload dan validasi tipe file.
10. **Riwayat Pembayaran di Admin** - Tampilkan riwayat pembayaran di dashboard admin untuk reconciliasi.

## Kontribusi
Jika Anda ingin berkontribusi, silakan fork repository ini dan ajukan pull request.

## Lisensi
Proyek ini adalah proyek internal untuk SDIT Insan Rabbani dan tidak dilisensikan untuk penggunaan komersial tanpa izin.