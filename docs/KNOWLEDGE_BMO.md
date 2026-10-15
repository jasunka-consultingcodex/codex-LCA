# Rekap Pareto BMO

Hasil ada di [Rekap_Pareto_BMO.xlsx](../Rekap_Pareto_BMO.xlsx), dibuat dari unggahan `SimaPro BMO.xlsx` (SHA-256 `d5ec2e10cb74f37dea3898da180813494226268b5678631d4ea637bf8115f25e`) dengan tampilan mengacu pada `01. Rekap_Pareto_LMO_GPT 2.xlsx` (SHA-256 `76e987edc6bc908d6709065deffb23a90849f7bf928bb970b94703fd60961d0e`). Kedua berkas masukan tidak disimpan di repo. Workbook hasil terdiri dari dua sheet: `Hotspot Proses BMO` berisi empat tabel PROC terpisah, dan `LCIA - LCI BMO Pareto` berisi tiga tabel LCI terpisah.

## Ketidaksesuaian sumber HULU

`PROC HULU BMO` memuat 20 blok kategori mentah. Empat belas blok bertipe `Process / Total of all processes`; setelah menggabungkan CEN dan CED, tersedia **10 kategori proses**. Enam blok lain — ODP, AP, EP, POFP, ABN, ABF — bertipe `Substance / Compartment / Total of all compartments`. Keenamnya **bukan data proses** dan tidak dicantumkan sebagai proses dominan. Kolomnya tetap terlihat dalam tabel PROC HULU, tetapi tanpa angka dan diberi keterangan tidak tersedia. Diperlukan ekspor PROC HULU yang benar untuk melengkapi keenam kategori itu.

Sebaliknya, `LCI HULU BMO` hanya memuat 14 blok mentah, yaitu 10 kategori setelah agregasi energi. Enam kategori inventori yang tersisip dalam `PROC HULU BMO` mempunyai struktur LCI dan tidak tumpang tindih dengan kategori LCI HULU. Keenamnya dipakai untuk melengkapi tabel LCI HULU menjadi 16 kategori, diberi tanda `†` pada kategori serta keterangan asal sheet. Ini adalah koreksi penempatan data yang transparan, bukan anggapan bahwa inventori tersebut adalah proses.

## Metode

Ekspor SimaPro menempatkan persentase di tabel kiri dan nilai dampak absolut di tabel kanan. Rekap memakai angka kanan, lalu memeriksa persentase kiri terhadap `100 × nilai / |total|` hingga toleransi 0,005 poin persentase. Per kategori, kontribusi searah hasil neto dihitung sebagai `nilai / total kategori`, diurutkan menurun, dan berhenti pada **baris pertama** dengan kumulatif `≥80%` sebelum pembulatan. Kontribusi bertanda yang ditampilkan adalah `nilai / |total kategori|`, seperti keluaran SimaPro. Untuk total negatif, persentase terpilih dapat negatif sementara kumulatif arah neto positif; nilai lebih dari 100% dapat muncul akibat aliran yang saling mengimbangi.

ETP terestrial, air tawar, dan laut tetap tiga kategori. CEN dan CED masing-masing menjumlahkan tiga subkategori MJ pada total dan pada identitas baris yang sama, baru kemudian menerapkan Pareto. Identitas proses adalah unit proses + nama proses; identitas LCI adalah unit proses + nama substansi + kompartemen. Angka dampak disimpan sebagai nilai Excel dengan format `0.00E+00`, persentase `0.00%`; koma atau titik desimal mengikuti pengaturan locale Excel.

Sheet LCI hanya menghubungkan inventori dengan **unit proses**. Nama proses yang ditampilkan berasal dari proses terbesar pada kategori dan unit yang sama di sheet PROC, sebagai konteks tingkat unit. Ini bukan bukti bahwa proses individual tersebut menghasilkan seluruh inventori. Untuk enam kategori HULU yang tidak punya data PROC, kolom proses individual menyatakan tidak tersedia.

| Sumber | Kategori hasil | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| PROC HULU BMO | 10 | 24 | 82,71% |
| PROC INTI BMO | 16 | 42 | 82,96% |
| PROC UT BMO | 16 | 38 | 81,01% |
| PROC DWN BMO | 16 | 33 | 81,32% |
| LCI HULU BMO + enam blok tersisip | 16 | 68 | 80,03% |
| LCI INTI BMO | 16 | 73 | 80,14% |
| LCI UT BMO | 16 | 68 | 80,66% |
| **Total** | **106** | **346** | |

Seluruh 106 kelompok yang memiliki data mencapai 80% dan berhenti pada baris pertama yang mencapainya. Tabel HULU PROC tetap belum lengkap untuk enam kategori yang tidak mempunyai data proses. Struktur dan angka workbook hasil diperiksa melalui `openpyxl`; tampilan PROC dan LCI diperiksa melalui ekspor PDF.

Generator: [scripts/generate_bmo_pareto.py](../scripts/generate_bmo_pareto.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_bmo_pareto.py 'SimaPro BMO.xlsx' '01. Rekap_Pareto_LMO_GPT 2.xlsx' Rekap_Pareto_BMO.xlsx
```
