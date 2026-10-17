# Rekap LMO setelah perbaikan GWP PROC HULU

Hasil terkini ialah [Rekap_LCI_LMO.xlsx](../Rekap_LCI_LMO.xlsx) (HULU, INTI, UT), [Rekap_PROC_LMO.xlsx](../Rekap_PROC_LMO.xlsx) (HULU, INTI, UT, DOWN), serta [Rekap_Pareto_LMO_Lengkap.xlsx](../Rekap_Pareto_LMO_Lengkap.xlsx) (ketujuh sheet dan `Audit Pareto`). Ketiganya mempertahankan format template yang dicatat pada [perbaikan LUC HULU sebelumnya](KNOWLEDGE_LMO_LUC_HULU.md). Workbook lima sheet tetap menjadi arsip.

| Masukan | SHA-256 |
| --- | --- |
| `SimaPro LMO.xlsx` (unggahan perbaikan GWP PROC HULU) | `71def352b8e106141481c69a590ad23e98a95a6e0cd60b19f321366cdb500580` |
| Unggahan sebelumnya setelah perbaikan LUC HULU | `88464094d19933f53964d404b4072c2000b1d7886d49c5a5b68f5025ce4f1322` |

## Perubahan GWP HULU

Sumber `PROC HULU LMO` kini memakai indikator `Characterization` dengan kategori `GWP100 - fossil`, sama seperti `LCI HULU LMO`. Sebelumnya PROC memakai `Damage assessment` / `GWP100`. Total PROC GWP turun dari `24,198492` menjadi `18,517792 kg CO2-eq`; total baru sama dengan LCI HULU. Beberapa nilai baris PROC juga berubah sesuai perubahan indikator. Hasil GWP LCI tidak berubah.

| Urutan Pareto | Unit proses / kegiatan pada PROC HULU | Nilai dampak | Kontribusi | Kumulatif |
| --- | --- | ---: | ---: | ---: |
| 1 | `5. Overburden Removal` / `5. Overburden Removal` | `14,445036` | 78,01% | 78,01% |
| 2 | `1. Supporting Material Production and Transportation` / `Ammonium nitrate {RoW}` (market for ammonium nitrate) | `1,3259296` | 7,16% | 85,17% |

Baris kedua diperlukan karena baris pertama belum mencapai 80%. Pada versi sebelumnya, baris kedua PROC ialah `Palm oil, crude {MY}` dari `5. Overburden Removal`. Pada LCI HULU, baris GWP tetap `Carbon dioxide, fossil` (Air) di `5. Overburden Removal`: `15,859743 kg CO2-eq`, kontribusi 85,65%.

Karena kategori dan total GWP kedua sheet kini cocok, kolom proses terkait pada LCI HULU kembali menampilkan `5. Overburden Removal`. Ini hanya **asosiasi pada kategori dan unit proses yang sama**. Sumber LCI tidak membuktikan bahwa proses individual itu adalah asal langsung dari inventori `Carbon dioxide, fossil`.

## Metode dan pemeriksaan

Perhitungan memakai nilai dampak pada tabel kanan ekspor SimaPro; persentase kiri diperiksa terhadapnya. Baris searah total neto diurutkan dari kontribusi terbesar sampai kumulatif pertama `≥80%`. ETP tetap tiga kategori terpisah; tiga subkategori CEN dan tiga subkategori CED masing-masing dijumlahkan sebelum Pareto. Nilai dampak berformat ilmiah `0.00E+00` dan kontribusi `0.00%` di Excel.

Ketujuh sheet masing-masing tetap memuat 16 kategori, seluruh **112 kelompok** mencapai ambang Pareto, dan total **347 baris** terpilih. Dibanding versi LUC sebelumnya, hanya nilai/identitas/kontribusi kelompok `PROC HULU LMO / GWP` yang berubah. Nomor baris sumber untuk kategori PROC HULU lain bergeser satu karena blok GWP pada ekspor baru lebih pendek. Sheet LCI HULU hanya berubah pada keterangan proses terkait GWP. Sheet INTI, UT, dan DOWN tidak berubah secara substantif. Workbook LCI dan PROC terpisah diperiksa terhadap sheet yang sama di workbook gabungan.

Generator: [scripts/generate_lmo_five_sheets.py](../scripts/generate_lmo_five_sheets.py). Sumber SimaPro dan dua template tidak disimpan di repo.
