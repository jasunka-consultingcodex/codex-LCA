# Catatan analisis Pareto LMO

Dokumen ini mencatat sumber, aturan hitung, hasil pemeriksaan, dan batas penafsiran untuk [Rekap_Pareto_LMO.xlsx](../Rekap_Pareto_LMO.xlsx). Hasil rinci per proses dan inventori, nilai dampak, persentase, kumulatif, serta nomor baris sumber berada di workbook tersebut. Catatan ini adalah dokumentasi hasil yang dibuat pada 14 Oktober 2026, bukan data LCA baru.

## Sumber dan cakupan

| Berkas masukan | Peran | SHA-256 |
| --- | --- | --- |
| `SimaPro LMO.xlsx` | Angka proses dan inventori dari `PROC HULU`, `PROC INTI`, `PROC UT`, `LCI HULU`, `LCI INTI`, `LCI UT` | `917c8125999539d8c3de6dc6d2d688f8aa88b809f9be41f36211fc839fbe029d` |
| `Analisis LCA Hotspot Process.xlsx` | Acuan tampilan `Hotspot LMO` | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx` | Acuan tampilan `LCIA - LCI LMO` | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Ketiga berkas masukan diunggah dalam percakapan dan **tidak** disimpan di repo. Acuan tampilan digunakan untuk susunan dan istilah, bukan sebagai sumber angka rekap. Workbook hasil berisi dua sheet: `Hotspot Proses LMO` dan `LCIA - LCI LMO Pareto`. Masing-masing mempunyai tabel HULU, INTI, dan UT yang terpisah. SHA-256 workbook hasil pada saat catatan ini dibuat: `875e55146966141269dcda64687fc38629e565e4d048b061717e63edb644c131`.

Satuan fungsional di ekspor SimaPro tertulis `1 p Upstream`, `1 p Core`, dan `1 p Utility`. Referensi visual membahas produksi 1 ton batu bara, tetapi angka ekspor tidak saya konversi ulang. Jangan menganggap total PROC dan LCI HULU GWP sebanding sebelum memeriksa metode kategorinya.

## Struktur sumber

Setiap sheet SimaPro tersusun dari blok berulang: metadata `Category`, tabel persentase di kiri, dan tabel nilai dampak absolut di kanan. Baris `Total of all processes` berlaku untuk PROC; `Total of all compartments` berlaku untuk LCI. Baris data menampilkan kontribusi per unit proses. Rekap memakai nilai absolut di tabel kanan untuk perhitungan, lalu menghitung ulang persentase dari total kategori. Dengan cara itu tiga subkategori energi dapat digabung dalam satuan yang sama tanpa menjumlahkan tiga kolom 100%.

Setiap fase dihitung terhadap total kategorinya sendiri. Persentase HULU, INTI, dan UT tidak dijumlahkan menjadi satu persentase lintas fase. Nama unit proses dan nama kegiatan/proses pada sheet PROC mengikuti label ekspor SimaPro; deskripsi kegiatan naratif tidak direka dari nama tersebut.

Unit proses per fase:

- HULU: Supporting Material Production and Transportation; Land Clearing; Top Soil Removal; Drilling and Blasting; Overburden Removal.
- INTI: Coal Mining; Coal Hauling; Coal Crushing; Coal Barging.
- UT: Reclamation; Wastewater Treatment Plant; Workshop; Mining Utility; Hazardous Waste Transportation; Solid Waste Transportation; Hazardous Waste Treatment System; Solid Waste Treatment System.

Ada 20 blok kategori unik pada PROC/LCI HULU, 20 pada PROC/LCI INTI, dan 19 pada LCI UT. `PROC UT` memuat 25 blok, karena keenam blok energi muncul dua kali. Awal pasangan blok yang berulang ialah 541/787, 582/828, 623/869, 664/910, 705/951, dan 746/992. Baris yang tersedia pada pengulangan cocok dengan blok pertama; blok pertama yang lengkap dipakai dan setiap kategori dihitung **sekali**. Sesudah penggabungan energi, ada 16 kategori hasil untuk HULU, 16 untuk INTI, dan 15 untuk UT. `Water use`/WSF tidak ada pada sumber UT, sehingga tidak dibuat baris WSF UT.

## Pemetaan kategori

| Hasil | Kategori SimaPro |
| --- | --- |
| GWP | `GWP100` atau `GWP100 - fossil`, mengikuti sheet sumber |
| ODP | `Ozone layer depletion (ODP)` |
| AP | `Acidification` |
| EP | `Eutrophication` |
| POFP | `Photochemical oxidation` |
| ABN | `Abiotic depletion` |
| ABF | `Abiotic depletion (fossil fuels)` |
| ETP Terrestrial | `Terrestrial ecotoxicity` |
| ETP Freshwater | `Freshwater ecotoxicity` |
| ETP Marine | `Marine ecotoxicity` |
| HTC | `Human carcinogenic toxicity` |
| HTN | `Human non-carcinogenic toxicity` |
| WSF | `Water use` |
| LUC | `Land use` |
| CEN | Jumlah `Non renewable, fossil` + `Non-renewable, nuclear` + `Non-renewable, biomass` |
| CED | Jumlah `Renewable, biomass` + `Renewable, wind, solar, geothe` + `Renewable, water` |

ETP **tidak** diagregasi. CEN dan CED masing-masing menjumlahkan tiga nilai dampak bersatuan MJ, baik pada total kategori maupun pada baris yang mempunyai identitas sama. Identitas baris PROC adalah **nama proses + unit proses**; identitas baris LCI adalah **nama substansi + kompartemen + unit proses**. Satu nama yang muncul pada dua unit proses tetap menjadi dua baris.

## Aturan Pareto

Untuk tiap fase dan kategori, kontribusi baris dihitung sebagai `nilai dampak baris / total dampak kategori`. Baris dengan kontribusi positif diurutkan menurun. Kumulatif dimulai dari baris terbesar dan berhenti tepat pada baris pertama dengan `kumulatif >= 80%` menurut angka lengkap, sebelum pembulatan tampilan. Karena penyebut adalah total **neto**, kontribusi individu atau kumulatif dapat melebihi 100% saat ada aliran lain yang berlawanan tanda; ini tampak jelas pada WSF. Aliran pengimbang yang menghasilkan kontribusi negatif tidak masuk urutan dominan. Nilai dampak tetap disimpan sebagai angka Excel, dengan format ilmiah `0.00E+00` (pemisah desimal mengikuti locale Excel); persentase memakai `0.00%`.

Pada sheet LCI, sumber hanya mengaitkan substansi dengan **unit proses**. Ekspor tidak menunjukkan pasangan langsung `substansi -> proses individual`. Kolom `Proses dominan terkait pada unit` diisi dari proses dengan kontribusi terbesar pada kategori dan unit proses yang sama di sheet PROC. Itu asosiasi tingkat unit, **bukan** bukti bahwa proses individual tersebut menghasilkan seluruh substansi pada baris LCI.

Perbedaan metode yang perlu diperhatikan: `PROC HULU` memakai kategori `GWP100` dengan indikator *Damage assessment*, sementara `LCI HULU` memakai `GWP100 - fossil` dengan indikator *Characterization*. Total GWP HULU keduanya memang berbeda. Jangan memaksa jumlahnya sama atau menafsirkan persentase keduanya sebagai basis identik.

## Penyimpanan dan akses

Workbook hasil dan dokumen ini ada pada branch GitHub `codex/rekap-pareto-lmo`; branch `main` belum memuatnya. Tautan `sandbox:` pada percakapan tidak dapat membaca berkas dari mesin cloud ini. Akses berkas melalui GitHub atau browser file lingkungan cloud. Jika analisis perlu dibuat ulang dari awal, sediakan lagi ketiga berkas masukan yang disebut di atas; berkas sumber tidak ikut dalam commit.

## Pemeriksaan hasil

| Fase | Kategori per sheet | Baris dominan PROC | Baris dominan LCI |
| --- | ---: | ---: | ---: |
| HULU | 16 | 35 | 57 |
| INTI | 16 | 40 | 75 |
| UT | 15 | 42 | 73 |
| **Total** | **47** | **117** | **205** |

Seluruh 47 kelompok pada **masing-masing** sheet mencapai ambang 80% dan berhenti pada baris pertama yang mencapainya. Pemeriksaan dilakukan pada workbook hasil yang dapat dibuka kembali oleh `openpyxl`. Kumulatif terakhir terendah adalah sekitar 80,26% pada PROC dan 80,00% pada LCI. Pada WSF LCI, baris pertama dapat melebihi 1000% atau 2400% terhadap total neto karena kompensasi aliran lain; nilainya tidak berarti penggunaan air bruto sebesar persentase itu.

Beberapa hasil untuk orientasi cepat (angka berikut adalah **kontribusi baris pertama**, bukan total kategori): HULU GWP dipimpin `5. Overburden Removal` (59,69%) pada PROC dan `Carbon dioxide, fossil` (85,65%) pada LCI. INTI GWP dipimpin `1. Coal Mining` (92,65%) pada PROC dan `Methane` (91,52%) pada LCI. INTI ABF dan CEN dipimpin `1. Coal Mining` serta `Coal, hard`, keduanya sekitar 99,78% dari hasil masing-masing. UT GWP dipimpin proses pengolahan `Waste mineral oil` dengan insinerasi limbah berbahaya (48,69%) pada PROC, sedangkan substansi LCI terbesar ialah `Carbon dioxide, fossil` (50,27%). Baca baris-baris berikutnya dalam workbook bila satu baris belum mencapai 80%.

## Jejak hasil per kategori

Tabel berikut adalah indeks pemeriksaan. `n` adalah jumlah baris yang dipilih sampai batas 80%; angka kumulatif adalah hasil **baris terakhir**. Workbook memuat nilai, nama, unit, dan nomor baris SimaPro untuk tiap baris dominan.

### PROSES HULU

| Kategori | Total PROC | Total LCI | n PROC | Kumulatif PROC | n LCI | Kumulatif LCI |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Potensi pemanasan global (GWP) | 2.42E+01 | 1.85E+01 | 2 | 85.21% | 1 | 85.65% |
| Penipisan ozon (ODP) | 2.01E-06 | 2.01E-06 | 1 | 87.96% | 1 | 90.08% |
| Pengasaman (AP) | 1.26E-01 | 1.26E-01 | 2 | 83.69% | 3 | 83.18% |
| Eutrofikasi (EP) | 2.93E-02 | 2.93E-02 | 2 | 84.10% | 4 | 83.14% |
| Oksidasi fotokimia (POFP) | 1.13E-02 | 1.13E-02 | 2 | 90.39% | 3 | 80.19% |
| Penipisan abiotik non-fosil (ABN) | 3.75E-05 | 3.75E-05 | 3 | 85.35% | 8 | 83.24% |
| Penipisan abiotik fosil (ABF) | 1.67E+02 | 1.67E+02 | 1 | 80.26% | 2 | 84.87% |
| Ekotoksisitas terestrial (ETP) | 1.50E+01 | 1.50E+01 | 3 | 88.74% | 5 | 82.22% |
| Ekotoksisitas air tawar (ETP) | 1.77E-02 | 1.77E-02 | 3 | 90.19% | 6 | 80.57% |
| Ekotoksisitas laut (ETP) | 2.38E-02 | 2.38E-02 | 3 | 84.23% | 10 | 80.07% |
| Toksisitas manusia karsinogenik (HTC) | 5.08E-02 | 5.08E-02 | 3 | 80.74% | 4 | 84.70% |
| Toksisitas manusia non-karsinogenik (HTN) | 3.99E-01 | 3.99E-01 | 2 | 111.49% | 3 | 90.49% |
| Kelangkaan air (WSF) | 3.23E+00 | 3.23E+00 | 3 | 95.68% | 1 | 1076.04% |
| Perubahan penggunaan lahan (LUC) | 4.93E+00 | 4.93E+00 | 2 | 92.94% | 3 | 86.00% |
| Energi tak terbarukan (CEN) | 1.85E+02 | 1.85E+02 | 2 | 83.74% | 2 | 81.66% |
| Energi terbarukan (CED) | 1.27E+02 | 1.27E+02 | 1 | 96.07% | 1 | 96.00% |

### PROSES INTI

| Kategori | Total PROC | Total LCI | n PROC | Kumulatif PROC | n LCI | Kumulatif LCI |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Potensi pemanasan global (GWP) | 4.50E+01 | 4.50E+01 | 1 | 92.65% | 1 | 91.52% |
| Penipisan ozon (ODP) | 5.20E-07 | 5.20E-07 | 2 | 81.65% | 2 | 82.04% |
| Pengasaman (AP) | 1.63E-02 | 1.63E-02 | 4 | 82.67% | 7 | 83.70% |
| Eutrofikasi (EP) | 3.61E-03 | 3.61E-03 | 4 | 81.62% | 7 | 82.50% |
| Oksidasi fotokimia (POFP) | 9.51E-03 | 9.51E-03 | 1 | 89.69% | 1 | 87.14% |
| Penipisan abiotik non-fosil (ABN) | 2.01E-06 | 2.01E-06 | 3 | 88.39% | 11 | 82.59% |
| Penipisan abiotik fosil (ABF) | 1.82E+04 | 1.82E+04 | 1 | 99.78% | 1 | 99.78% |
| Ekotoksisitas terestrial (ETP) | 1.74E+00 | 1.74E+00 | 4 | 86.56% | 9 | 82.73% |
| Ekotoksisitas air tawar (ETP) | 1.88E-03 | 1.88E-03 | 4 | 89.54% | 8 | 81.04% |
| Ekotoksisitas laut (ETP) | 3.73E-03 | 3.73E-03 | 4 | 87.96% | 12 | 80.67% |
| Toksisitas manusia karsinogenik (HTC) | 5.25E-03 | 5.25E-03 | 4 | 87.05% | 5 | 81.14% |
| Toksisitas manusia non-karsinogenik (HTN) | 3.66E-02 | 3.66E-02 | 1 | 100.22% | 3 | 85.20% |
| Kelangkaan air (WSF) | 9.19E-02 | 9.19E-02 | 2 | 86.30% | 1 | 2438.08% |
| Perubahan penggunaan lahan (LUC) | 3.77E-01 | 3.77E-01 | 2 | 88.66% | 4 | 80.18% |
| Energi tak terbarukan (CEN) | 1.91E+04 | 1.91E+04 | 1 | 99.77% | 1 | 99.78% |
| Energi terbarukan (CED) | 1.26E+01 | 1.26E+01 | 2 | 97.83% | 2 | 97.76% |

### PROSES UTILITAS (UT)

| Kategori | Total PROC | Total LCI | n PROC | Kumulatif PROC | n LCI | Kumulatif LCI |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Potensi pemanasan global (GWP) | 6.58E-01 | 6.58E-01 | 2 | 84.16% | 2 | 89.22% |
| Penipisan ozon (ODP) | 3.92E-08 | 3.92E-08 | 2 | 80.97% | 2 | 83.05% |
| Pengasaman (AP) | 2.43E-03 | 2.43E-03 | 4 | 83.28% | 6 | 80.52% |
| Eutrofikasi (EP) | 5.39E-04 | 5.39E-04 | 4 | 82.83% | 6 | 81.18% |
| Oksidasi fotokimia (POFP) | 2.10E-04 | 2.10E-04 | 3 | 81.99% | 7 | 80.48% |
| Penipisan abiotik non-fosil (ABN) | 4.18E-07 | 4.18E-07 | 4 | 84.53% | 11 | 80.00% |
| Penipisan abiotik fosil (ABF) | 3.10E+00 | 3.10E+00 | 3 | 84.67% | 3 | 82.19% |
| Ekotoksisitas terestrial (ETP) | 5.30E-01 | 5.30E-01 | 3 | 87.76% | 6 | 81.03% |
| Ekotoksisitas air tawar (ETP) | 6.13E-04 | 6.13E-04 | 3 | 83.93% | 6 | 82.44% |
| Ekotoksisitas laut (ETP) | 1.03E-03 | 1.03E-03 | 3 | 81.86% | 9 | 80.33% |
| Toksisitas manusia karsinogenik (HTC) | 1.66E-03 | 1.66E-03 | 4 | 88.52% | 4 | 82.16% |
| Toksisitas manusia non-karsinogenik (HTN) | 2.80E-02 | 2.80E-02 | 2 | 97.12% | 3 | 84.33% |
| Perubahan penggunaan lahan (LUC) | 6.46E-02 | 6.46E-02 | 1 | 82.91% | 3 | 82.41% |
| Energi tak terbarukan (CEN) | 3.41E+00 | 3.41E+00 | 3 | 81.84% | 4 | 84.78% |
| Energi terbarukan (CED) | 2.25E+00 | 2.25E+00 | 1 | 87.34% | 1 | 87.27% |
