# InputBPB — Input Data BPB/BTB MAPPI

Aplikasi Streamlit untuk membaca lembar **Biaya Teknis Bangunan (BTB)** MAPPI —
dari PDF cetakan portal `member.mappi.or.id` maupun dari XLSX/CSV — lalu
menyimpan dan mengunduhnya kembali dalam format **XLSX**.

```bash
pip install -r requirements.txt
streamlit run input5.py
```

## Mengapa versi sebelumnya menghasilkan tabel kosong

Cetakan PDF dari halaman *BTB Interaktif* **tidak memiliki lapisan teks sama
sekali**: seluruh huruf digambar sebagai kurva vektor (halaman contoh berisi 0
karakter, 0 font, dan 4.135 kurva). Akibatnya
`pdfplumber.extract_tables()` selalu mengembalikan daftar kosong, sehingga
`extract_data_from_pdf()` mengembalikan `[]` dan seluruh alur penyimpanan
bekerja di atas tabel kosong.

Garis-garis tabelnya tetap ada sebagai objek vektor. Modul `btb_io.py`
memanfaatkan hal itu: grid sel direkonstruksi dari garis tersebut, lalu isi tiap
sel dibaca satu per satu — memakai lapisan teks bila ada, atau OCR bila tidak.
Membaca per sel jauh lebih akurat daripada OCR satu halaman penuh, yang pada
berkas contoh menghilangkan baris berwarna abu-abu (`PEMBULATAN`,
`TOTAL BIAYA PEMBANGUNAN BARU`) dan beberapa angka.

## Yang diperbaiki

**Pembacaan**

- Rekonstruksi grid tabel dari garis vektor PDF, lalu pembacaan per sel
  (lapisan teks bila tersedia, OCR bila tidak) — seluruh 21 baris × 11 kolom
  terbaca dari berkas contoh.
- Metadata **Provinsi**, **Kota/Kabupaten**, dan **Tahun** terisi otomatis dari
  dokumen; sebelumnya harus diketik manual.
- Nama baris dan kolom dicocokkan ke istilah baku formulir BTB, sehingga salah
  baca kecil pada OCR tidak merusak struktur tabel.
- Tambahan sumber data **XLSX/XLSM/CSV**, dalam bentuk tabel maupun bentuk
  panjang.
- Angka format Indonesia (`1.780.811`) maupun Inggris (`1,780,811`) diuraikan
  dengan benar. Sebelumnya semua koma dibuang tanpa memeriksa perannya.
- `pd.to_numeric(..., errors="ignore")` dihapus — opsi tersebut sudah dihapus
  dari pandas 2.2+ dan membuat aplikasi gagal pada pandas versi baru.

**Validasi**

- Konsistensi aritmatika diperiksa per kolom: jumlah komponen = `TOTAL ( A )`,
  `TOTAL ( B )`, `A + B`, `PPN 11%`, dan pembulatan. Ketidaksesuaian
  ditampilkan per sel, dan baris total dapat dihitung ulang.
- Hasil bacaan dapat dikoreksi langsung di layar sebelum disimpan.

**Penulisan XLSX**

- Angka ditulis sebagai bilangan asli dengan format ribuan `#,##0`, bukan teks;
  kepala tabel berwarna, panel dibekukan, lebar kolom menyesuaikan, dan filter
  otomatis aktif.
- Menyimpan ulang kota/tahun yang sama **menimpa** data lama alih-alih
  menggandakannya.
- Kepala tabel tidak lagi hilang saat menyimpan: sebelumnya `DataFrame`
  dibongkar ke `values.tolist()` sehingga nama kolom berganti angka 0…N.
- Unduhan tersedia dalam tiga bentuk: **tabel**, **transpose**, dan **panjang**
  (lihat di bawah). Fungsi transpose lama menghapus kolom elemen dan
  menghasilkan berkas tanpa keterangan baris.
- Tombol unduh tidak lagi bersarang di dalam tombol lain — pada versi lama
  tombol unduh langsung hilang begitu halaman dimuat ulang.

## Bentuk berkas XLSX

Tiga bentuk tersedia pada tab **Download Data**, dan ketiganya juga ditulis
sebagai sheet terpisah pada basis data (`BTB Data`, `BTB Long`,
`BTB Transpose`). Ketiganya dapat diunggah kembali ke aplikasi.

**Tabel** — satu baris per elemen biaya, seperti dokumen aslinya:

| Provinsi | Kota/Kabupaten | Tahun | No | Kelompok | Elemen | Mewah | … |
| --- | --- | --- | --- | --- | --- | --- | --- |
| PROV. DKI JAKARTA | KOTA JAKARTA | 2026 | 3 | A. BIAYA LANGSUNG | Pondasi | 941.613 | … |

**Transpose** — keterangan turun di kolom pertama; setiap elemen biaya menjadi
satu kolom, sehingga beberapa kota/tahun bisa berjajar dalam satu lembar:

| | | | | |
| --- | --- | --- | --- | --- |
| **Provinsi** | PROV. DKI JAKARTA | PROV. DKI JAKARTA | PROV. DKI JAKARTA | … |
| **Kota/Kabupaten** | KOTA JAKARTA | KOTA JAKARTA | KOTA JAKARTA | … |
| **Tahun** | 2026 | 2026 | 2026 | … |
| **No** | 1 | 2 | 3 | … |
| **Kelompok** | A. BIAYA LANGSUNG | A. BIAYA LANGSUNG | A. BIAYA LANGSUNG | … |
| **Elemen** | A. BIAYA LANGSUNG | Persiapan | Pondasi | … |
| **Mewah** | | 0 | 941.613 | … |
| **Menengah** | | 0 | 803.399 | … |
| … | | | | |

**Panjang** — satu baris per kombinasi elemen × tipe bangunan, siap dijadikan
PivotTable.

## Berkas

| Berkas | Isi |
| --- | --- |
| `input5.py` | Aplikasi Streamlit (3 tab: Input, Data Telah Diinput, Download) |
| `btb_io.py` | Logika baca/tulis BTB, tanpa Streamlit |
| `test_btb_io.py` | 39 uji, termasuk uji terhadap PDF contoh |
| `samples/` | Lembar BTB contoh (KOTA JAKARTA, 2026) |
| `requirements.txt` | Dependensi Python |
| `packages.txt` | Paket sistem untuk Streamlit Cloud (`tesseract-ocr`) |

## Menjalankan uji

```bash
pip install pytest
pytest test_btb_io.py -v
```

Uji yang membutuhkan OCR otomatis dilewati bila `tesseract` belum terpasang.

## Catatan penyimpanan

Basis data disimpan di `/tmp/data_btb/btb_data.xlsx`. Di Streamlit Cloud isi
`/tmp` terhapus saat aplikasi tidur, jadi **unduh berkasnya** setelah menginput.
Bila dijalankan di server sendiri, arahkan ke folder permanen:

```bash
export BTB_DATA_DIR=/data/btb
```

## OCR

OCR hanya diperlukan untuk PDF tanpa lapisan teks. Di luar Streamlit Cloud,
pasang tesseract secara manual:

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-ind   # Debian/Ubuntu
brew install tesseract tesseract-lang                  # macOS
```

Bila tesseract tidak tersedia, aplikasi tetap berjalan dan menampilkan pesan
yang menyarankan unggahan XLSX/CSV.
