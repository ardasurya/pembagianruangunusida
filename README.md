# Optimasi Pembagian Ruangan Kuliah UNUSIDA

Aplikasi Streamlit untuk optimasi ruangan berdasarkan data jadwal kuliah dan master ruang UNUSIDA.

## Master ruang bawaan
Aplikasi sudah menyertakan `DATA RUANG.xlsx` dengan 25 ruang:
- 16 ruang kapasitas 25 mahasiswa
- 9 ruang kapasitas 40 mahasiswa

Master ruang dapat diganti dari sidebar jika ada perubahan data.

## Aturan optimasi
- Kelas <=25 mahasiswa diprioritaskan ke ruang kapasitas 25.
- Jika ruang kecil penuh, kelas <=25 dapat memakai ruang besar jika opsi diaktifkan.
- Kelas 26-40 mahasiswa memakai ruang kapasitas 40.
- Kelas >40 mahasiswa ditandai `PERLU SPLIT`.
- Ruang yang sama tidak boleh dipakai oleh dua kelas yang waktunya bertumpuk pada hari yang sama.

## Dashboard okupansi 12 prodi
Aplikasi selalu menampilkan 12 prodi UNUSIDA:
1. Informatika
2. Sistem Informasi
3. DKV
4. Manajemen
5. Akuntansi
6. Teknik Lingkungan
7. Teknik Kimia
8. Teknik Industri
9. Pendidikan Guru Sekolah Dasar
10. Pendidikan Bahasa Inggris
11. Pendidikan Guru Madrasah Ibtidaiyah
12. Pendidikan Islam Anak Usia Dini

### Definisi indikator
- **Okupansi Prodi (%)** = total mahasiswa pada jadwal teralokasi / total kapasitas kursi ruang yang dipakai prodi x 100%.
- **Keberhasilan Alokasi (%)** = jadwal teralokasi / total jadwal prodi x 100%.
- **Porsi Waktu Ruang (%)** = total menit penggunaan ruang prodi / seluruh menit penggunaan ruang x 100%.

## Menjalankan
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Output Excel
Hasil download mencakup sheet:
- Hasil Optimasi
- Ringkasan
- Okupansi Prodi
- Penggunaan Ruang
- Master Ruang
- Konflik (jika ditemukan)
