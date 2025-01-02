# Sertifikat Seminar Otomatis

skrip otomatis untuk mempermudah dalam membuat sertifikat peserta dan mengirimkannya secara otomatis melalui email. Dengan alat ini, Anda hanya perlu memberikan daftar peserta dalam format Excel dan template sertifikat, maka sistem akan menghasilkan sertifikat PDF dan mengirimkannya ke email peserta.

## Persyaratan
Pastikan Anda telah menginstal:
- Python 3.8 atau lebih baru.
- Pustaka Python berikut:
  ```bash
  pip install pandas docxtpl docx2pdf python-dotenv
  ```
- Sistem operasi Windows untuk fitur konversi PDF (karena menggunakan `docx2pdf`).

## Instalasi
1. **Clone repositori ini**:
   ```bash
   git clone https://github.com/username/sertifikat-seminar.git
   cd sertifikat-seminar
   ```

2. **Konfigurasikan file `.env`**:
   Buat file `.env` di direktori utama dan tambahkan konfigurasi email:
   ```env
   EMAIL_ADDRESS=your-email@gmail.com
   EMAIL_PASSWORD=your-app-password
   SMTP_SERVER=smtp.gmail.com
   SMTP_PORT=465
   ```

   **Catatan:** Jika menggunakan Gmail, aktifkan **2-Step Verification** dan gunakan **App Password**.

3. **Siapkan file pendukung**:
   - File Excel dengan data peserta (contoh: `pendaftaran.xlsx`) bisa di simpan dalam folder form_sertifikat.
   - Template sertifikat dalam format Word (.docx) **dengan placeholder `{{NAMA}}` untuk nama peserta**.

## Cara Penggunaan
1. instal semua pustaka / liblary yang diperlukan
2. Jalankan skrip dengan perintah berikut:
   ```bash
   python buatSertifikat.py
   ```

3. Masukkan input sesuai panduan yang muncul, seperti:
   - Jalur file Excel peserta.
   - Jalur template sertifikat.
   - Kolom nama dan email pada file Excel.
   - Subjek dan isi pesan email.

4. Skrip akan:
   - Membuat sertifikat untuk setiap peserta.
   - Mengirimkan sertifikat ke email peserta.

## Struktur Direktori
- `buatSertifikat.py`: Skrip utama untuk membuat dan mengirim sertifikat.
- `sendemail.py`: Modul untuk pengiriman email.
- `.env`: File konfigurasi untuk email.
- `sertifikat/`: Direktori keluaran untuk file PDF sertifikat.
- `form_sertifikat/`: Direktori untuk menaruh file excel.
- `template_sertifikat/`: digunakan untuk menaruh file template dokumen

## Contoh Input File
### Excel (pendaftaran.xlsx):
| Nama Peserta         | Email                |
|----------------------|----------------------|
| Agus Sunardi         | agus@gmail.com       |
| Faisal fajar         | faisal@gmail.com     |

### Template Sertifikat (template.docx):
Template harus memiliki placeholder `{{NAMA}}` di posisi nama peserta.

## FAQ
### 1. Bagaimana jika saya tidak menggunakan Windows?
Fitur konversi PDF menggunakan `docx2pdf` hanya mendukung Windows.

### 2. Mengapa email tidak terkirim?
- Pastikan kredensial email di `.env` benar.
- Periksa koneksi internet.
- Cek apakah port SMTP (465 atau 587) diblokir oleh jaringan Anda => jika anda menggunakan jaringan intansi, sekolah, perusahaan disarankan menggunakan jaringan hostpot atau lainnya.

### 3. Bagaimana bila terjadi masalah yang tudak di duga?
Anda bisa membaca ulanga persyaratan dan penggunaan kode ini. jika masih belum teratasi anda bisa membuat isu di repositori ini, terkait masalah yang anda hadapi.


>Terimakasih telah mampir pada Repositori ini semoga bermanfaat dan memudahkan pekerjaan yang berulang saat membuat sertifikat
