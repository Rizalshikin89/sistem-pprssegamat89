# Sistem PPRS – versi pembangunan setempat

## Jalankan di Windows
1. Pasang Python 3.11 atau lebih baharu.
2. Pasang LibreOffice (untuk penukaran PDF).
3. Ekstrak folder `sistem-pprs` ke komputer.
4. Buka Command Prompt dalam folder itu.
5. Jalankan:
   ```bat
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   py app.py
   ```
6. Buka `http://127.0.0.1:5000`.

## Penting
- Templat di `templates/` ialah sumber asal. Jangan edit atau gantikan fail tersebut.
- Output dijana sebagai salinan di `output/`.
- Aplikasi ini mengikat beberapa medan yang dikenal pasti kepada ruangan dalam dokumen. Ia masih versi pembangunan; semak setiap dokumen dengan teliti sebelum pengesahan rasmi.
- Lampiran Q1 ialah fail `.doc` lama dan buat masa ini disalin tanpa pengisian automatik.
- Bahagian tandatangan / pengesahan pihak luar tidak diisi secara automatik.
- PDF memerlukan LibreOffice yang boleh dipanggil melalui arahan `libreoffice`.
- Jalankan hanya pada komputer dipercayai. Jangan hoskan tanpa kawalan akses di GitHub Pages atau laman awam kerana data penerima ialah maklumat peribadi.
