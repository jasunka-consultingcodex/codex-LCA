# Rekap ulang LMO setelah WSF proses INTI dilengkapi

Hasil gabungan ada di [Rekap_Pareto_LMO_5_Sheet.xlsx](../Rekap_Pareto_LMO_5_Sheet.xlsx). Lima sheet hasilnya ialah `LCI INTI LMO`, `PROC INTI LMO`, `LCI UT LMO`, `PROC UT LMO`, dan `PROC DOWN LMO`; sheet keenam `Audit Pareto` mencatat nilai, kontribusi, kumulatif, dan baris sumber. Dua berkas dengan format sama tersedia terpisah: [Rekap_LCI_LMO.xlsx](../Rekap_LCI_LMO.xlsx) berisi dua sheet LCI, sedangkan [Rekap_PROC_LMO.xlsx](../Rekap_PROC_LMO.xlsx) berisi tiga sheet PROC.

| Masukan | Fungsi | SHA-256 |
| --- | --- | --- |
| `SimaPro LMO.xlsx` (unggahan perbaikan WSF) | Seluruh angka rekap | `d4755a5aa2eeb5bbff8faef96f02af9701d09834ed63f535da17ba9cb81afdc3` |
| `Analisis LCA Hotspot Process.xlsx`, sheet `Hotspot LMO` | Acuan tata letak PROC | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx`, sheet `LCIA - LCI LMO` | Acuan tata letak LCI | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Ketiga berkas masukan diunggah dalam percakapan dan tidak disimpan di repo. Sheet hilir pada sumber bernama `PROC DWN LMO`, sedangkan hasil diberi nama `PROC DOWN LMO`. Hasil baru dibuat dari angka pada unggahan perbaikan, bukan dari rekap lama.

## Metode dan batas atribusi

Setiap blok SimaPro menempatkan persentase di kiri dan nilai dampak absolut di kanan. Rekap memakai nilai kanan dan memeriksa persentase kiri terhadap `100 × nilai / |total|`. Identitas PROC ialah unit proses + nama proses; identitas LCI ialah unit proses + nama inventori + kompartemen. Nama yang sama pada unit atau kompartemen berbeda tetap dihitung sebagai baris berbeda.

Untuk tiap kategori, kontribusi searah total neto dihitung sebagai `nilai / total kategori`, diurutkan menurun, dan pemilihan berhenti **pada baris pertama** ketika kumulatif mencapai `≥80%` sebelum pembulatan. Kontribusi bertanda yang ditampilkan ialah `nilai / |total kategori|`, mengikuti SimaPro. Jika total neto negatif, persentase terpilih dapat negatif sementara kumulatif arah neto tetap positif. Kontribusi yang melebihi 100% mungkin terjadi akibat aliran pengimbang. Nilai dampak tetap berupa angka Excel dengan format `0.00E+00`, persentase `0.00%`; pemisah desimal mengikuti locale Excel.

ETP terestrial, air tawar, dan laut tetap terpisah. CEN dan CED masing-masing menjumlahkan **tiga subkategori** bersatuan MJ pada total dan identitas baris yang sama sebelum Pareto dihitung. Kelima sheet target kini memiliki 16 kategori hasil.

Sumber LCI menghubungkan inventori dengan **unit proses**, tanpa hubungan langsung ke proses individual. Proses pada kolom LCI ialah proses dengan kontribusi searah neto terbesar pada kategori dan unit proses yang sama di PROC. Itu hanya konteks tingkat unit. Misalnya, inventori dominan LCI INTI WSF berada pada `4. Coal Barging` dan proses terkait pada unit itu ialah proses diesel. Ini tidak membuktikan bahwa proses diesel menghasilkan seluruh inventori `Water, turbine use` tersebut.

## Perbaikan WSF INTI dan pemeriksaan

Unggahan terdahulu tidak memiliki WSF di `PROC INTI LMO`, sehingga rekap PROC INTI hanya memuat 15 kategori dan kolom proses terkait pada LCI INTI WSF tidak dapat diisi. Unggahan perbaikan memiliki blok `Water use` bertipe proses. Total WSF PROC INTI adalah `9,1838854E-02 m³`; Pareto memilih proses `Palm oil, crude {MY}| palm oil mill operation | Cut-off, S` pada `2. Coal Hauling` (56,20%), lalu proses yang sama pada `1. Coal Mining` (30,16%), sehingga kumulatifnya 86,36%.

| Sheet hasil | Kategori | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| LCI INTI LMO | 16 | 74 | 80,12% |
| PROC INTI LMO | 16 | 40 | 81,66% |
| LCI UT LMO | 16 | 69 | 80,00% |
| PROC UT LMO | 16 | 40 | 80,68% |
| PROC DOWN LMO | 16 | 33 | 81,32% |
| **Total** | **80** | **256** | |

Seluruh 80 kelompok mencapai 80% dan setiap baris sebelumnya masih di bawah 80% menurut angka penuh. Perbandingan dengan workbook sebelumnya menunjukkan **hanya satu kelompok Pareto baru**, yaitu PROC INTI WSF; seluruh nilai pada 79 kelompok lama tetap sama. Kolom proses terkait pada LCI INTI WSF kini terisi dari kategori dan unit proses yang cocok. Workbook dapat dibuka ulang dengan `openpyxl`, dan tampilan hasil diperiksa melalui ekspor PDF.

Generator: [scripts/generate_lmo_five_sheets.py](../scripts/generate_lmo_five_sheets.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_lmo_five_sheets.py 'SimaPro LMO.xlsx' 'Analisis LCA Hotspot Process.xlsx' 'Analisis LCA dan LCI terhadap LCIA.xlsx' Rekap_Pareto_LMO_5_Sheet.xlsx --lci-output Rekap_LCI_LMO.xlsx --proc-output Rekap_PROC_LMO.xlsx
```
