# Rekap Pareto GMO

Hasil tersedia dalam [Rekap_LCI_GMO.xlsx](../Rekap_LCI_GMO.xlsx) dengan sheet HULU, INTI, UT; [Rekap_PROC_GMO.xlsx](../Rekap_PROC_GMO.xlsx) dengan sheet HULU, INTI, UT, DWN; dan [Rekap_Pareto_GMO_Lengkap.xlsx](../Rekap_Pareto_GMO_Lengkap.xlsx) dengan ketujuh sheet tersebut serta `Audit Pareto`. **LCI DWN tidak dibuat**, sesuai permintaan pengguna dan karena sumber mentahnya tidak disertakan. Format LCI mengikuti `Analisis LCA dan LCI terhadap LCIA`; format proses mengikuti `Analisis LCA Hotspot Process`.

| Berkas masukan | Peran | SHA-256 |
| --- | --- | --- |
| `SimaPro GMO.xlsx` | Nilai dampak proses dan inventori GMO | `c442ed629020d72b6d3866ccb635d12763582384f3963ada276d37fe8d7167c6` |
| `Analisis LCA Hotspot Process.xlsx`, sheet `Hotspot LMO` | Acuan tata letak PROC | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx`, sheet `LCIA - LCI LMO` | Acuan tata letak LCI | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Berkas masukan yang diunggah tidak disimpan di repo. Sumber menyediakan tiga sheet LCI dan empat sheet PROC yang dibutuhkan. Kategori dan total PROC/LCI cocok untuk HULU, INTI, dan UT, sehingga proses pada sheet LCI dapat dicantumkan sebagai **asosiasi kategori dan unit proses**. Ekspor LCI tidak menghubungkan inventori langsung ke proses individual; nama proses terkait bukan bukti asal langsung inventori tersebut.

## Metode dan pemeriksaan

Perhitungan memakai nilai dampak pada tabel kanan ekspor SimaPro. Persentase tabel kiri diperiksa terhadap `100 × nilai / |total|` dengan toleransi 0,005 poin persentase. Identitas baris PROC ialah unit proses + nama proses; identitas baris LCI ialah unit proses + nama inventori + kompartemen. Baris searah total neto diurutkan dari kontribusi terbesar sampai baris pertama yang membuat kumulatif `≥80%` sebelum pembulatan.

ETP terestrial, air tawar, dan laut tetap tiga kategori. CEN dan CED masing-masing menjumlahkan tiga subkategori MJ pada total dan baris beridentitas sama sebelum pemilihan Pareto. Angka dampak disimpan sebagai angka Excel berformat `0.00E+00`; persentase berformat `0.00%`. Total WSF neto negatif pada UT dan PROC DWN; kontribusi bertanda di workbook bisa negatif dan besarnya bisa melebihi 100% karena aliran pengimbang. Kumulatif arah neto dan nomor baris sumber tersedia pada sheet audit.

| Sheet sumber | Kategori | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| LCI HULU GMO | 16 | 57 | 80,76% |
| PROC HULU GMO | 16 | 31 | 80,06% |
| LCI INTI GMO | 16 | 78 | 80,28% |
| PROC INTI GMO | 16 | 43 | 80,55% |
| LCI UT GMO | 16 | 68 | 80,92% |
| PROC UT GMO | 16 | 39 | 80,07% |
| PROC DWN GMO | 16 | 33 | 81,32% |
| **Total** | **112** | **349** | |

Seluruh 112 kelompok mencapai ambang Pareto dan berhenti pada baris pertama yang mencapainya. Setiap nilai audit dicocokkan dengan baris dan total sumber; sheet workbook terpisah dibandingkan sel per sel dengan workbook gabungan. Tata letak LCI dan PROC diperiksa melalui ekspor PDF.

Generator: [scripts/generate_gmo_pareto.py](../scripts/generate_gmo_pareto.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_gmo_pareto.py 'SimaPro GMO.xlsx' 'Analisis LCA Hotspot Process.xlsx' 'Analisis LCA dan LCI terhadap LCIA.xlsx' Rekap_Pareto_GMO_Lengkap.xlsx --lci-output Rekap_LCI_GMO.xlsx --proc-output Rekap_PROC_GMO.xlsx
```
