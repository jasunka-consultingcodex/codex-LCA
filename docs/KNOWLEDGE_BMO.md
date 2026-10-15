# Rekap Pareto BMO (sumber HULU telah diperbaiki)

Hasil terbaru ada di [Rekap_Pareto_BMO.xlsx](../Rekap_Pareto_BMO.xlsx). Workbook berisi dua sheet: `Hotspot Proses BMO` dengan empat tabel terpisah (HULU, INTI, UT, DWN) dan `LCIA - LCI BMO Pareto` dengan tiga tabel terpisah (HULU, INTI, UT).

| Berkas masukan | Fungsi | SHA-256 |
| --- | --- | --- |
| `SimaPro BMO.xlsx` (unggahan perbaikan) | Semua angka proses dan inventori BMO | `e60b2b08672a31007c6b803470a475f9acf8688bcecb22e60aa6ee8adcdc4fc4` |
| `01. Rekap_Pareto_LMO_GPT 2.xlsx` | Acuan format workbook | `76e987edc6bc908d6709065deffb23a90849f7bf928bb970b94703fd60961d0e` |

Berkas masukan diunggah dalam percakapan dan tidak disimpan di repo. Unggahan BMO terdahulu memiliki enam blok `Substance/Compartment` yang tersisip pada `PROC HULU BMO`, sementara `LCI HULU BMO` kekurangan enam kategori itu. **Unggahan terbaru memperbaiki keduanya:** `PROC HULU BMO` kini memiliki 20 blok bertipe `Process`, dan `LCI HULU BMO` memiliki 20 blok bertipe `Substance`. Karena itu, catatan keterbatasan HULU lama dan tanda `†` pada workbook lama tidak lagi berlaku. Tidak ada data silang sheet yang diperlukan untuk rekap baru. Sumber hilir tetap bernama `PROC DWN BMO`.

## Metode

Setiap blok SimaPro berisi persentase di tabel kiri dan nilai dampak absolut di tabel kanan. Rekap memakai angka kanan; persentase kiri diperiksa terhadap `100 × nilai / |total|` dengan toleransi 0,005 poin persentase. Untuk tiap kategori, kontribusi searah hasil neto dihitung sebagai `nilai / total kategori`, diurutkan menurun, dan dipilih sampai **baris pertama** dengan kumulatif `≥80%` sebelum pembulatan. Persentase yang ditampilkan adalah `nilai / |total kategori|`, seperti di SimaPro. Pada kategori dengan total neto negatif, persentase terpilih dapat negatif sedangkan kumulatif arah neto positif. Persentase di atas 100% dapat muncul akibat aliran pengimbang.

ETP terestrial, air tawar, dan laut tetap tiga kategori. CEN dan CED masing-masing menjumlahkan tiga subkategori bersatuan MJ pada total dan identitas baris yang sama, lalu dihitung Pareto. Identitas baris PROC adalah unit proses + nama proses; identitas LCI adalah unit proses + nama inventori + kompartemen. Nilai dampak adalah angka Excel berformat `0.00E+00`; kontribusi dan kumulatif berformat `0.00%`. Pemisah desimal mengikuti pengaturan locale Excel.

Sumber LCI hanya menghubungkan inventori dengan **unit proses**. Nama proses dalam tabel LCI ialah proses dengan kontribusi terbesar pada kategori dan unit yang sama di sheet PROC, sebagai konteks tingkat unit. Ini bukan bukti bahwa proses individual tersebut menghasilkan seluruh inventori pada baris LCI.

| Sumber | Kategori hasil | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| PROC HULU BMO | 16 | 37 | 80,49% |
| PROC INTI BMO | 16 | 42 | 82,96% |
| PROC UT BMO | 16 | 38 | 81,01% |
| PROC DWN BMO | 16 | 33 | 81,32% |
| LCI HULU BMO | 16 | 68 | 80,03% |
| LCI INTI BMO | 16 | 73 | 80,14% |
| LCI UT BMO | 16 | 68 | 80,66% |
| **Total** | **112** | **359** | |

Seluruh 112 kelompok mencapai ambang 80% dan berhenti pada baris pertama yang mencapainya. Ketujuh tabel mempunyai 16 kategori setelah agregasi energi. Struktur dan angka workbook diperiksa melalui `openpyxl`; tampilan PROC dan LCI diperiksa melalui ekspor PDF.

Generator: [scripts/generate_bmo_pareto.py](../scripts/generate_bmo_pareto.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_bmo_pareto.py 'SimaPro BMO.xlsx' '01. Rekap_Pareto_LMO_GPT 2.xlsx' Rekap_Pareto_BMO.xlsx
```
