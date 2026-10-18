# Rekap Pareto SMO

Hasil berada di [Rekap_LCI_SMO.xlsx](../Rekap_LCI_SMO.xlsx), [Rekap_PROC_SMO.xlsx](../Rekap_PROC_SMO.xlsx), dan [Rekap_Pareto_SMO_Lengkap.xlsx](../Rekap_Pareto_SMO_Lengkap.xlsx). Workbook LCI mengikuti format `Analisis LCA dan LCI terhadap LCIA`; workbook PROC mengikuti format `Analisis LCA Hotspot Process`. Workbook gabungan memuat delapan sheet fase/jenis serta `Audit Pareto`.

| Berkas masukan | Peran | SHA-256 |
| --- | --- | --- |
| `SimaPro SMO.xlsx` | Nilai dampak SMO | `2c83a1e42be047c308f18a1936e2b2258c2ada78328d8db3c67799607c425ce0` |
| `Analisis LCA Hotspot Process.xlsx`, sheet `Hotspot LMO` | Acuan tata letak PROC | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx`, sheet `LCIA - LCI LMO` | Acuan tata letak LCI | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Masukan yang diunggah tidak disimpan di repo. Sumber berisi `LCI HULU SMO`, `LCI INTI SMO`, `LCI UT SMO`, serta `PROC HULU SMO`, `PROC INTI SMO`, `PROC UT SMO`, `PROC DWN SMO`. **Tidak ada sheet `LCI DWN SMO`.** Karena inventori hilir tidak bisa diturunkan dari tabel proses atau total LCIA, `LCI DWN SMO` dalam workbook hasil hanya menampilkan keterangan bahwa datanya belum tersedia. Tidak ada angka inventori hilir yang diisi. `PROC DWN SMO` tetap dihitung dari sumbernya.

## Metode

Setiap kategori dihitung dari nilai dampak pada tabel kanan ekspor SimaPro; persentase di tabel kiri dicocokkan terhadap `100 × nilai / |total|` dengan toleransi 0,005 poin persentase. Identitas proses adalah unit proses + nama proses. Identitas LCI adalah unit proses + nama substansi + kompartemen. Baris yang kontribusinya searah total neto diurutkan menurun dan diambil sampai **baris pertama** yang mencapai kumulatif `≥80%` sebelum pembulatan.

ETP terestrial, air tawar, dan laut tetap tiga kategori. CEN dan CED masing-masing menggabungkan tiga subkategori MJ pada total dan baris beridentitas sama sebelum Pareto. Nilai dampak disimpan sebagai angka Excel berformat `0.00E+00`; kontribusi dan kumulatif berformat `0.00%`. Tanda desimal mengikuti locale Excel. Total neto WSF bernilai negatif pada UT dan PROC DWN; persentase bertanda dapat negatif dan kontribusi terhadap total neto dapat melebihi 100% akibat aliran pengimbang.

Proses yang disebut pada sheet LCI ialah **proses dominan terkait pada kategori dan unit yang sama** di sheet PROC. Sumber LCI tidak membuktikan hubungan langsung inventori → proses individual. Pada tiga fase yang memiliki kedua jenis sheet, label kategori dan total PROC/LCI sesuai sebelum asosiasi tingkat unit ditampilkan.

| Sheet sumber | Kategori | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| LCI HULU SMO | 16 | 58 | 80,50% |
| PROC HULU SMO | 16 | 34 | 80,47% |
| LCI INTI SMO | 16 | 75 | 80,14% |
| PROC INTI SMO | 16 | 40 | 80,25% |
| LCI UT SMO | 16 | 80 | 80,11% |
| PROC UT SMO | 16 | 46 | 80,15% |
| PROC DWN SMO | 16 | 33 | 81,32% |
| **Total tersedia** | **112** | **366** | |

Seluruh 112 kelompok tersedia berhenti pada baris pertama yang mencapai ambang 80%. Nilai dan format angka pada workbook terpisah dibandingkan sel per sel dengan sheet gabungan. Tampilan LCI, PROC, serta penanda LCI DWN diperiksa melalui ekspor PDF. Sheet `Audit Pareto` mencatat nomor baris sumber untuk setiap kontribusi dominan.

Generator: [scripts/generate_smo_pareto.py](../scripts/generate_smo_pareto.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_smo_pareto.py 'SimaPro SMO.xlsx' 'Analisis LCA Hotspot Process.xlsx' 'Analisis LCA dan LCI terhadap LCIA.xlsx' Rekap_Pareto_SMO_Lengkap.xlsx --lci-output Rekap_LCI_SMO.xlsx --proc-output Rekap_PROC_SMO.xlsx
```
