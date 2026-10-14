# Catatan rekap proses hilir LMO

[Hotspot_LMO_PROC_DWN.xlsx](../Hotspot_LMO_PROC_DWN.xlsx) merangkum proses dominan pada fase hilir dari unggahan SimaPro LMO yang baru. Pengguna menyebut sheet `PROC DOWN`; nama yang ada di workbook adalah **`PROC DWN`**. Produk sumber tertulis `1 p Downstream (of project 2026_PT Berau Coal LMO)`.

## Sumber dan tampilan

| Berkas | Peran | SHA-256 |
| --- | --- | --- |
| Unggahan baru `SimaPro LMO.xlsx` | Angka `PROC DWN` | `f3138c9ce70bc5ec8837e51d166f4b89e2721c00d099d95d77da210e5f58df24` |
| `Analisis LCA Hotspot Process.xlsx` | Template visual sheet `Hotspot LMO` | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |

File hasil memiliki sheet pertama `Hotspot PROC DOWN` dengan susunan horizontal seperti template: **Unit Proses**, **Proses**, satu kolom per dampak, nilai dampak pada baris atas dan kontribusi pada baris bawah. Warna, border, gaya header, dan susunan baris diambil dari `Hotspot LMO`. Dua kolom ETP tambahan dibuat agar terestrial, air tawar, dan laut tetap terpisah. Satuan di header mengikuti ekspor SimaPro, tanpa mengubah `1 p Downstream` menjadi satuan per ton. Sheet kedua `Audit Pareto` mencatat urutan, kumulatif, dan nomor baris sumber.

## Perhitungan

Metode umum mengikuti [catatan Pareto LMO](KNOWLEDGE_LMO.md). Setiap blok SimaPro mempunyai persentase di kiri dan nilai dampak di kanan. Rekap memakai nilai dampak kanan. Persentase **bertanda** yang ditampilkan mengikuti SimaPro: `nilai proses pada unit / |total dampak kategori|`. Untuk memilih Pareto, kontribusi **searah total neto** dihitung sebagai `nilai proses pada unit / total dampak kategori`; kandidat positif diurutkan menurun, lalu daftar berhenti pada baris pertama yang mencapai sedikitnya 80% dengan presisi penuh. Identitas baris adalah **unit proses + nama proses**. Generator membandingkan persentase yang dihitung dengan tabel kiri SimaPro untuk mendeteksi salah baca kolom.

Ada 20 kategori sumber yang menjadi 16 hasil. CEN menjumlahkan `Non renewable, fossil`, `Non-renewable, nuclear`, dan `Non-renewable, biomass`. CED menjumlahkan `Renewable, biomass`, `Renewable, wind, solar, geothe`, dan `Renewable, water`. Ketiganya pada masing-masing kelompok bersatuan MJ. ETP terestrial, air tawar, dan laut tetap tiga kolom. `PROC DWN` tidak mempunyai blok kategori ganda. Urutan kolom mengikuti template, dengan ABF sebelum ABN. Format angka dampak adalah `0.00E+00`; format kontribusi adalah `0.00%`.

Unit proses sumber: Coal Transport Domestic; Coal Transport Export; Coal Mill; Boiler PLTU; Turbine PLTU; End of Life FABA. Dalam pilihan Pareto 80%, lima pasangan proses-unit muncul pada empat unit: Coal Transport Export, Coal Mill, Boiler PLTU, dan Turbine PLTU. Coal Transport Domestic serta End of Life FABA **tidak masuk pilihan**; itu tidak berarti dampaknya nol. Sheet pertama memuat lima pasangan tersebut dalam tata letak template, sementara sheet audit memuat 33 baris terpilih lintas 16 kategori.

`Water use`/WSF mempunyai total neto negatif `-6.2488108E-02 m3`. Baris `5. Turbine PLTU` bernilai `-1.3377858E-01 m3`, sehingga persentase bertanda sesuai SimaPro adalah **−214,09%**. Kumulatif Pareto searah total neto menjadi **214,09%** dan berhenti pada baris itu. Besarnya melebihi 100% karena proses lain mengimbangi dampak tersebut; angka itu bukan persentase dari total bruto.

## Pemeriksaan per kategori

`n` adalah jumlah baris sampai batas 80%; kumulatif adalah nilai pada baris terakhir. Seluruh 16 kategori memenuhi ambang dan berhenti pada baris pertama yang mencapainya.

| Kategori | Total dampak | n | Kumulatif |
| --- | ---: | ---: | ---: |
| GWP | 3.57E-02 | 2 | 81,32% |
| ODP | 2.19E-09 | 2 | 83,89% |
| AP | 8.01E-04 | 2 | 82,92% |
| EP | 9.65E-05 | 3 | 92,09% |
| POFP | 2.05E-05 | 3 | 92,28% |
| ABF | 3.85E-01 | 2 | 82,16% |
| ABN | 2.56E-08 | 2 | 85,15% |
| ETP Terrestrial | 4.13E-02 | 2 | 85,08% |
| ETP Freshwater | 1.65E-05 | 2 | 82,45% |
| ETP Marine | 5.00E-05 | 2 | 83,84% |
| HTC | 3.27E-04 | 2 | 83,87% |
| HTN | 4.40E-03 | 2 | 84,04% |
| WSF | -6.25E-02 | 1 | 214,09% |
| LUC | 2.30E-04 | 2 | 83,85% |
| CEN | 4.25E-01 | 2 | 82,11% |
| CED | 1.09E-02 | 2 | 90,05% |

GWP didominasi listrik tegangan tinggi pada `4. Boiler PLTU` (55,02%) dan angkutan laut barang curah pada `2. Coal Transport Export` (26,30%). AP dipimpin `4. Boiler PLTU` (55,21%); EP juga dipimpin proses itu (50,75%). Nilai lengkap dan kategori lain ada di workbook hasil.

## Pembuatan ulang

Berkas sumber yang diunggah tidak disimpan dalam repo. Setelah kedua berkas sumber tersedia lokal, jalankan:

```bash
python scripts/generate_proc_dwn.py 'SimaPro LMO.xlsx' 'Analisis LCA Hotspot Process.xlsx' Hotspot_LMO_PROC_DWN.xlsx
```

Generator memeriksa 16 kategori hasil, kecocokan tabel persentase dengan nilai dampak sumber, total dan satuan yang valid, serta tercapainya ambang 80% pada setiap kategori. SHA-256 workbook hasil: `a0c21baaa9a92695625145c43c36ef255f4982488106078b3c56c76f64d452c4`.
