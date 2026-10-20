# Rekap Pareto BSSR

Hasil ada di [Rekap_LCI_BSSR.xlsx](../Rekap_LCI_BSSR.xlsx), [Rekap_PROC_BSSR.xlsx](../Rekap_PROC_BSSR.xlsx), dan [Rekap_Pareto_BSSR_Lengkap.xlsx](../Rekap_Pareto_BSSR_Lengkap.xlsx). Masing-masing jenis rekap memuat HULU, INTI, dan UT. Workbook gabungan juga memiliki `Audit Pareto`. Format LCI mengikuti `Analisis LCA dan LCI terhadap LCIA`; format proses mengikuti `Analisis LCA Hotspot Process`.

| Berkas masukan | Peran | SHA-256 |
| --- | --- | --- |
| `SimaPro BSSR.xlsx` | Nilai dampak proses dan inventori BSSR | `65595c67f02211668963d3c91c8daf8eb038e4fdedec21e93bb1feeb61407cbe` |
| `Analisis LCA Hotspot Process.xlsx`, sheet `Hotspot LMO` | Acuan tata letak PROC | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx`, sheet `LCIA - LCI LMO` | Acuan tata letak LCI | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Berkas masukan yang diunggah tidak disimpan di repo. Nama unit proses BSSR mengikuti ekspor SimaPro tanpa awalan angka; nama tersebut dipertahankan. Sumber memuat keenam sheet HULU, INTI, UT yang diminta.

## Blok WSF UT yang tersisip

`LCI UT BSSR` tidak memuat kategori `Water use`. Di `PROC UT BSSR`, terdapat **dua** blok `Water use` berturut-turut: blok pertama mulai baris 299 dengan header `Substance` dan total `Total of all compartments` (inventori); blok kedua mulai baris 333 dengan header `Process` dan total `Total of all processes` (proses). Keduanya memakai kategori dan unit proses yang sama serta total neto `−1,8327431 m3`. Rekap LCI UT mengambil **hanya blok inventori pertama** dari file sumber itu. Rekap PROC UT mengambil **hanya blok proses kedua**. Tidak ada nilai yang ditaksir atau digandakan.

Baris Pareto WSF LCI UT adalah `Settling Pond` / inventori `Waste water/m3` di kompartemen `Water`, bernilai `−7,1348152 m3`. Kontribusi bertanda terhadap `|total|` ialah `−389,30%`, sedangkan kumulatif searah total neto ialah `389,30%`; persentase di atas 100% muncul karena aliran air lain mengimbangi nilai neto tersebut. Di `Audit Pareto`, sumber baris ini ditulis eksplisit sebagai `PROC UT BSSR!310`. Sheet LCI UT juga mencantumkan catatan sumber khusus.

## Metode dan pemeriksaan

Rekap mengambil nilai dampak dari tabel kanan setiap blok SimaPro; persentase tabel kiri diperiksa terhadap `100 × nilai / |total|` dengan toleransi 0,005 poin persentase. Identitas baris PROC adalah unit proses + nama proses; identitas baris LCI adalah unit proses + inventori + kompartemen. Baris searah total neto diurutkan dari kontribusi terbesar, lalu dipilih hingga **baris pertama** yang membuat kumulatif mencapai `≥80%` menurut angka penuh.

ETP terestrial, air tawar, dan laut tetap terpisah. CEN dan CED masing-masing menjumlahkan tiga subkategori MJ pada total dan identitas baris yang sama sebelum Pareto. Dampak disimpan sebagai angka Excel berformat `0.00E+00`, kontribusi dan kumulatif `0.00%` (tanda desimal mengikuti locale Excel). Nama proses pada sheet LCI ialah proses dominan di kategori dan unit proses yang sama pada sheet PROC, **bukan bukti asal langsung** inventori. Seluruh kategori dan total PROC/LCI cocok di ketiga fase setelah blok WSF ditempatkan menurut tipe tabelnya.

| Sheet hasil | Kategori | Baris Pareto | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| LCI HULU BSSR | 16 | 57 | 80,88% |
| PROC HULU BSSR | 16 | 41 | 82,75% |
| LCI INTI BSSR | 16 | 84 | 80,28% |
| PROC INTI BSSR | 16 | 35 | 81,28% |
| LCI UT BSSR | 16 | 74 | 80,17% |
| PROC UT BSSR | 16 | 36 | 80,84% |
| **Total** | **96** | **327** | |

Seluruh 96 kelompok mencapai ambang Pareto dan berhenti pada baris pertama yang mencapainya. Nilai audit dibandingkan dengan total dan baris sumber; kedua workbook terpisah dibandingkan sel per sel dengan workbook gabungan. Tata letak LCI dan PROC diperiksa melalui ekspor PDF.

Generator: [scripts/generate_bssr_pareto.py](../scripts/generate_bssr_pareto.py). Jalankan dari akar repo dengan `openpyxl` terpasang:

```bash
python scripts/generate_bssr_pareto.py 'SimaPro BSSR.xlsx' 'Analisis LCA Hotspot Process.xlsx' 'Analisis LCA dan LCI terhadap LCIA.xlsx' Rekap_Pareto_BSSR_Lengkap.xlsx --lci-output Rekap_LCI_BSSR.xlsx --proc-output Rekap_PROC_BSSR.xlsx
```
