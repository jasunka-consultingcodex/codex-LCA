# Rekap ulang lima sheet LMO (unggahan terbaru)

Hasil terbaru ada di [Rekap_Pareto_LMO_5_Sheet.xlsx](../Rekap_Pareto_LMO_5_Sheet.xlsx). Workbook ini memakai **unggahan SimaPro terbaru** dan tidak mengubah workbook hasil sebelumnya. Sheet hasil: `LCI INTI LMO`, `PROC INTI LMO`, `LCI UT LMO`, `PROC UT LMO`, dan `PROC DOWN LMO`; sheet keenam `Audit Pareto` memberi jejak nilai, persentase, kumulatif, dan baris sumber. Berkas sumber yang diunggah tidak disimpan di GitHub.

| Masukan | Fungsi | SHA-256 |
| --- | --- | --- |
| `SimaPro LMO.xlsx` (unggahan terbaru) | Seluruh angka rekap | `63c18c10e0c2d53c6b36e279ddca91d1e195f273ae2cd96e6678d42328d4c7aa` |
| `Analisis LCA Hotspot Process.xlsx`, sheet `Hotspot LMO` | Acuan tata letak PROC | `87cd52484d4f9a25448866dcf30028b079669014feab5f3ecfb6e9999c108994` |
| `Analisis LCA dan LCI terhadap LCIA.xlsx`, sheet `LCIA - LCI LMO` | Acuan tata letak LCI | `04137ba0744323764ab46c7bb826435472f17a9c69c02faaa1031d51466e89e6` |

Nama sheet hilir pada sumber ialah `PROC DWN LMO`, sedangkan sheet hasil diberi nama `PROC DOWN LMO` sesuai permintaan. Setiap sheet sumber terdiri dari blok kategori dengan persentase di kiri dan nilai dampak absolut di kanan. Rekap dihitung dari **nilai kanan**; persentase kiri diperiksa terhadap nilai/|total|. Identitas proses adalah pasangan unit proses + nama proses. Identitas LCI adalah unit proses + nama inventori + kompartemen. Nama sama pada unit/kompartemen berbeda tetap dibedakan.

Untuk tiap kategori, baris searah total neto diurutkan menurut `nilai baris / total kategori`, terbesar lebih dulu. Pemilihan berhenti **tepat pada baris pertama** saat kumulatif mencapai 80%. Kolom kontribusi bertanda memakai `nilai baris / |total kategori|`, sebagaimana persentase SimaPro; pada total neto negatif persentase terpilih bisa negatif, sementara kumulatif arah neto tetap positif. Persentase di atas 100% dimungkinkan oleh aliran yang saling mengimbangi. Nilai dampak adalah angka Excel berformat `0.00E+00` dan persentase `0.00%`; separator desimal ditentukan locale Excel.

ETP terestrial, air tawar, dan laut tetap tiga kategori. CEN dan CED masing-masing menjumlahkan **tiga subkategori** bersatuan MJ pada total dan pada identitas baris yang sama, lalu Pareto dihitung dari agregat. `PROC INTI LMO` pada unggahan ini tidak memuat WSF sehingga kolom PROC INTI diberi tanda tidak tersedia dan kategori itu tidak direkayasa. Sebaliknya, `LCI INTI LMO`, `LCI UT LMO`, `PROC UT LMO`, dan `PROC DWN LMO` memuat WSF. Ini berbeda dari unggahan sebelumnya; angka lama tidak dipakai ulang.

Sheet LCI mengaitkan inventori dengan **unit proses**, tanpa hubungan langsung ke proses individual. Proses pada kolom LCI ialah proses dengan kontribusi searah neto terbesar pada kategori dan unit yang sama di sheet PROC. Itu penunjuk konteks unit, bukan atribusi inventori ke proses individual. Khusus LCI INTI WSF, sumber PROC INTI tidak punya kategori WSF, sehingga kolom itu menyatakan proses individual tidak tersedia.

| Sheet hasil | Kategori setelah agregasi | Baris terpilih | Kumulatif akhir terendah |
| --- | ---: | ---: | ---: |
| LCI INTI LMO | 16 | 74 | 80,12% |
| PROC INTI LMO | 15 | 38 | 81,66% |
| LCI UT LMO | 16 | 69 | 80,00% |
| PROC UT LMO | 16 | 40 | 80,68% |
| PROC DOWN LMO | 16 | 33 | 81,32% |
| **Total** | **79** | **254** | |

Seluruh 79 kelompok mencapai 80%, dan setiap baris sebelumnya masih di bawah 80% menurut angka penuh. Tampilan kedua jenis sheet diperiksa melalui ekspor PDF; workbook dapat dibuka ulang dengan `openpyxl`.

Generator: [scripts/generate_lmo_five_sheets.py](../scripts/generate_lmo_five_sheets.py). Jalankan dari akar repo dengan:

```bash
python scripts/generate_lmo_five_sheets.py 'SimaPro LMO.xlsx' 'Analisis LCA Hotspot Process.xlsx' 'Analisis LCA dan LCI terhadap LCIA.xlsx' Rekap_Pareto_LMO_5_Sheet.xlsx
```

Ketiga masukan perlu disediakan ulang secara lokal; jangan menambahkan berkas unggahan mentah ke repo tanpa kebutuhan khusus. Generator memakai `openpyxl` dan mengimpor pemetaan kategori dari `scripts/generate_proc_dwn.py`.
