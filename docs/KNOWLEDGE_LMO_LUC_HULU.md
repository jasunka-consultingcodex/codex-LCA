# Rekap LMO setelah perbaikan LUC HULU

Hasil terkini tersedia dalam tiga workbook: [Rekap_LCI_LMO.xlsx](../Rekap_LCI_LMO.xlsx) berisi HULU, INTI, UT; [Rekap_PROC_LMO.xlsx](../Rekap_PROC_LMO.xlsx) berisi HULU, INTI, UT, DOWN; dan [Rekap_Pareto_LMO_Lengkap.xlsx](../Rekap_Pareto_LMO_Lengkap.xlsx) berisi ketujuh sheet tersebut serta `Audit Pareto`. Workbook lima sheet [Rekap_Pareto_LMO_5_Sheet.xlsx](../Rekap_Pareto_LMO_5_Sheet.xlsx) ialah hasil sebelumnya tanpa HULU dan dipertahankan sebagai arsip.

| Masukan | Fungsi | SHA-256 |
| --- | --- | --- |
| `SimaPro LMO.xlsx` (unggahan perbaikan LUC HULU) | Nilai dampak PROC dan LCI | `88464094d19933f53964d404b4072c2000b1d7886d49c5a5b68f5025ce4f1322` |
| `Analisis LCA Hotspot Process.xlsx`, sheet `Hotspot LMO` | Acuan tata letak PROC | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx`, sheet `LCIA - LCI LMO` | Acuan tata letak LCI | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Berkas sumber yang diunggah tidak disimpan di repo. Sheet hilir sumber bernama `PROC DWN LMO` dan ditampilkan sebagai `PROC DOWN LMO` dalam hasil. Angka rekap dihitung dari tabel nilai dampak di sisi kanan tiap blok ekspor SimaPro; persentase di sisi kiri diperiksa terhadap nilai tersebut. Identitas proses adalah unit proses + nama proses; identitas LCI adalah unit proses + nama inventori + kompartemen.

## Perubahan LUC HULU

Dibanding unggahan sebelumnya, nilai sumber yang berubah dalam tujuh sheet target hanya kategori `Land use` pada `PROC HULU LMO` dan `LCI HULU LMO`. Total LUC keduanya naik dari `4,9319554` menjadi `12,380609 m²a crop eq`. Sheet `LCIA LMO` dan `UNC LMO` pada sumber juga berubah; rekap ini mengambil angka langsung dari PROC dan LCI, bukan dari kedua sheet itu. Semua kelompok kategori INTI, UT, dan DOWN tetap identik dengan hasil sebelumnya.

| Sumber | Baris Pareto baru, berurutan | Kumulatif akhir |
| --- | --- | ---: |
| PROC HULU LUC | `2. Land Clearing` / proses `2. Land Clearing`: `8,7171566` (70,41%); `5. Overburden Removal` / proses `Palm oil, crude {MY}`: `3,3151813` (26,78%) | 97,19% |
| LCI HULU LUC | `2. Land Clearing` / `Transformation, from forest, intensive, ID` (Raw): `8,5475982` (69,04%); `5. Overburden Removal` / `Occupation, permanent crop` (Raw): `1,8451383` (14,90%) | 83,94% |

Nama proses di kolom LCI adalah **asosiasi tingkat unit**: proses dengan kontribusi terbesar pada kategori dan unit proses yang sama di PROC. Ekspor LCI tidak memberikan pasangan inventori → proses individual. Header LCI kini menyatakan bahwa proses terkait **bukan asal inventori yang terverifikasi**. Khusus GWP HULU, kategori sumber PROC ialah `GWP100` dan LCI `GWP100 - fossil`; proses individual sengaja tidak dipadankan karena dasar indikator keduanya berbeda.

## Aturan dan pemeriksaan

Pareto dihitung per kategori dari nilai baris dibagi total dampak kategori. Baris searah total neto diurutkan menurun, lalu dipilih sampai **baris pertama** yang membuat kumulatif mencapai `≥80%` menurut angka penuh. Persentase kontribusi bertanda yang ditampilkan ialah `nilai / |total|`; kumulatif arah neto ada di workbook gabungan. Pada total neto negatif, persentase bertanda dapat negatif. ETP terestrial, air tawar, dan laut tetap tiga kategori. CEN dan CED masing-masing menjumlahkan tiga subkategori MJ sebelum pemilihan Pareto. Angka dampak disimpan dalam format Excel `0.00E+00`, persentase `0.00%` (tanda desimal mengikuti locale Excel).

| Sheet | Kategori | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| LCI HULU LMO | 16 | 56 | 80,07% |
| PROC HULU LMO | 16 | 35 | 80,26% |
| LCI INTI LMO | 16 | 74 | 80,12% |
| PROC INTI LMO | 16 | 40 | 81,66% |
| LCI UT LMO | 16 | 69 | 80,00% |
| PROC UT LMO | 16 | 40 | 80,68% |
| PROC DOWN LMO | 16 | 33 | 81,32% |
| **Total** | **112** | **347** | |

Seluruh 112 kelompok mencapai ambang 80% dan berhenti pada baris pertama yang mencapainya. Seluruh 80 kelompok dari workbook lima sheet sebelumnya mempunyai baris audit yang identik. Workbook terpisah dibandingkan sel per sel dengan sheet pada workbook gabungan. Tata letak kedua jenis sheet juga diperiksa melalui ekspor PDF.

Generator yang diperbarui: [scripts/generate_lmo_five_sheets.py](../scripts/generate_lmo_five_sheets.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_lmo_five_sheets.py 'SimaPro LMO.xlsx' 'Analisis LCA Hotspot Process.xlsx' 'Analisis LCA dan LCI terhadap LCIA.xlsx' Rekap_Pareto_LMO_Lengkap.xlsx --lci-output Rekap_LCI_LMO.xlsx --proc-output Rekap_PROC_LMO.xlsx
```
