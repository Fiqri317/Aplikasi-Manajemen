
# Aplikasi Manajemen

Proyek ini adalah aplikasi manajemen yang dibangun menggunakan VB.NET, yang terdiri dari beberapa form yang dirancang untuk berbagai keperluan, seperti manajemen data supplier, konsumen, dokter, obat, dan transaksi pembelian. Aplikasi ini juga dilengkapi dengan form login dan menu utama untuk navigasi.

## Daftar Form

- **Form Login**: Form untuk masuk ke dalam aplikasi.
- **Form Menu**: Form utama yang digunakan sebagai navigasi ke form-form lainnya.
- **Form Data Supplier**: Form untuk mengelola data supplier.
- **Form Data Konsumen**: Form untuk mengelola data konsumen.
- **Form Data Dokter**: Form untuk mengelola data dokter.
- **Form Data Obat**: Form untuk mengelola data obat.
- **Form Transaksi Pembelian**: Form untuk mengelola transaksi pembelian.

## Prasyarat

Pastikan Anda sudah memiliki:

- Visual Studio yang mendukung pengembangan VB.NET.
- Database **Quiz.mdb** yang akan Anda sesuaikan lokasinya setelah meng-clone proyek ini.

## Instalasi dan Penggunaan

1. **Clone Repository**  
   Clone repository ini dengan menggunakan perintah berikut:
   ```bash
   git clone https://github.com/username/repository.git
   ```

2. **Buka Proyek di Visual Studio**  
   Setelah clone selesai, buka proyek ini di Visual Studio.

3. **Update Lokasi Database**  
   Sebelum menjalankan aplikasi, Anda perlu mengganti path `Data Source` untuk database agar sesuai dengan lokasi direktori database di komputer Anda. Path database yang perlu diubah adalah:
   
   ```plaintext
   Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Quiz\database\Quiz.mdb
   ```
   
   Sesuaikan path ini agar mengarah ke lokasi **Quiz.mdb** pada komputer Anda.

4. **Menjalankan Aplikasi**  
   Setelah mengganti sumber database, Anda dapat menjalankan aplikasi dengan menekan tombol **Start** di Visual Studio.

## Penggunaan Aplikasi

- **Login**: Masukkan username dan password untuk login.
- **Navigasi**: Gunakan form menu untuk mengakses form data supplier, konsumen, dokter, obat, atau transaksi pembelian.
  
## Struktur Proyek

- `FormLogin.vb` - Form Login
- `FormMenu.vb` - Form Menu Utama
- `FormDataSupplier.vb` - Form Data Supplier
- `FormDataKonsumen.vb` - Form Data Konsumen
- `FormDataDokter.vb` - Form Data Dokter
- `FormDataObat.vb` - Form Data Obat
- `FormTransaksiPembelian.vb` - Form Transaksi Pembelian

---

Silakan hubungi saya jika ada pertanyaan lebih lanjut atau ada masalah dengan pengaturan database.
