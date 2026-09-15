# Optimasi Pembagian Ruangan UNUSIDA - Embedded

Data jadwal default dari `Data Kelas 261 fixxxxxxxxxxx.xlsx` sudah di-embed langsung ke `app.py`.

Repository minimum:
- app.py
- requirements.txt

Tidak perlu menyertakan file jadwal Excel atau DATA RUANG.xlsx saat deploy.
User tetap dapat meng-upload Excel baru untuk mengganti data default pada sesi tersebut.

Jalankan:
```bash
pip install -r requirements.txt
streamlit run app.py
```
