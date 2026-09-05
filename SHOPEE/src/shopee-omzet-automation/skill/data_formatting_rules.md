# Aturan Format & Transformasi Data Laporan Mingguan ShopeeFood

Dokumen ini menjelaskan semua aturan transformasi dan *formatting* yang diterapkan pada data laporan mentah (*raw data*) ShopeeFood sebelum digabungkan menjadi file `Master_Weekly_Report.xlsx`. Seluruh logika ini diimplementasikan di Phase 3 pada skrip `weekly/run_weekly.py`.

## 1. Penambahan Kolom Identitas Merchant
- **`Merchant Name`**: Secara otomatis ditambahkan di kolom paling kiri (Index 0). Ini krusial karena saat belasan laporan toko digabungkan menjadi satu file Master, kolom ini menjadi satu-satunya penanda transaksi tersebut berasal dari cabang/outlet mana.

## 2. Standarisasi Kolom Nominal Uang (Harga, Diskon, Voucher)
Karena pengaturan *locale* pada Excel/sistem, nominal ribuan dari ShopeeFood seringkali terbaca sebagai desimal oleh Pandas (contoh: `12.500` terbaca sebagai `12.5`). Oleh karena itu:
- **Dikali 1000**: Nilai asli dikalikan 1000 untuk mengembalikan digit ribuan yang hilang.
- **Konversi murni ke Integer**: Nilai tersebut dipaksa menjadi bilangan bulat murni (`.astype(int)`).
- **Tanpa Simbol "Rp" & Tanpa Titik**: Nilai disimpan sebagai angka mentah (raw number). Tujuannya agar saat dibuka di Excel, kolom tersebut benar-benar dikenali sebagai format *Number*, bukan format *Text*.

**Daftar kolom yang terkena aturan ini:**
`Harga Makanan`, `Diskon`, `Diskon Flash Sale`, `Biaya Tambahan`, `Subsidi Merchant untuk Voucher Deals`, `Subsidi Platform untuk Flash Sale`, `Subsidi Voucher Makanan`, `Diskon Langsung`, `Nilai Transaksi`, `Harga Checkout Murah`.

## 3. Perhitungan Kolom Finansial Tambahan
Tiga metrik bisnis baru disisipkan secara dinamis berdasarkan data nominal yang sudah dikalibrasi di langkah sebelumnya:
- **`Commission`**: Dihitung dari `25% * Nilai Transaksi`. Berbeda dengan harga makanan yang dibulatkan, nilai komisi **mempertahankan angka desimalnya** (jika ada, contoh: `14500.75`), sesuai dengan potongan matematis riil.
- **`Revenue`**: `Nilai Transaksi - Commission`.
- **`OFD Fees`**: `Harga Makanan - Revenue`.

## 4. Pencegahan Scientific Notation pada Nomor Pesanan
- **`No. Pesanan`**: Nomor resi ShopeeFood bisa mencapai belasan digit (contoh: `3076998123811329188`). Jika dibiarkan sebagai format *Number*, Excel seringkali memotong presisinya atau menampilkannya dalam format eksponensial (`3.07E+18`).
- **Solusi**: Skrip secara paksa mengubah *tipe data* kolom ini menjadi `String/Text`, dan membuang akhiran desimal `.0` bawaan Pandas jika ada.

## 5. Standarisasi Format Tanggal Internasional
- **`Waktu Penyelesaian`**: Data mentah ShopeeFood mengekspor waktu penyelesaian menggunakan format tanggal dan penulisan bulan Indonesia (contoh: `"07 Mei 2026 23:16"`).
- **Transformasi**: Skrip memuat sebuah kamus *translator* bulan Indonesia-Inggris (Januari -> January, dll) untuk membantu Pandas memahami tanggal tersebut.
- **Hasil Akhir**: Tanggal ditulis ulang menjadi format yang sangat rapi dan konsisten: `"YYYY-MM-DD at HH:MM"` (contoh: `"2026-05-07 at 23:16"`). Jika baris gagal diterjemahkan (karena anomali data), ia akan aman dikembalikan ke bentuk aslinya.
