# Konteks proyek

Untuk pekerjaan terkait rekap LMO, baca [docs/KNOWLEDGE_LMO.md](docs/KNOWLEDGE_LMO.md) sebelum mengubah analisis atau menafsirkan workbook. Dokumen tersebut mencatat sumber, aturan Pareto, pemetaan kategori, duplikasi data, dan batas atribusi LCI. Hasil rinci ada di [Rekap_Pareto_LMO.xlsx](Rekap_Pareto_LMO.xlsx).

Untuk proses hilir, baca juga [docs/KNOWLEDGE_PROC_DWN.md](docs/KNOWLEDGE_PROC_DWN.md). Sheet SimaPro bernama `PROC DWN`; hasilnya ada di [Hotspot_LMO_PROC_DWN.xlsx](Hotspot_LMO_PROC_DWN.xlsx), dan generatornya di [scripts/generate_proc_dwn.py](scripts/generate_proc_dwn.py).

Untuk arsip perbaikan WSF INTI sebelum LUC HULU diperbarui, baca [docs/KNOWLEDGE_LMO_5_SHEET.md](docs/KNOWLEDGE_LMO_5_SHEET.md). Workbook lima sheet INTI, UT, DOWN dari saat itu berada di [Rekap_Pareto_LMO_5_Sheet.xlsx](Rekap_Pareto_LMO_5_Sheet.xlsx). Generator ada di [scripts/generate_lmo_five_sheets.py](scripts/generate_lmo_five_sheets.py). Sheet sumber hilir bernama `PROC DWN LMO`, sedangkan sheet hasil `PROC DOWN LMO`.

Untuk hasil LMO paling baru setelah perbaikan GWP PROC HULU, baca [docs/KNOWLEDGE_LMO_GWP_HULU.md](docs/KNOWLEDGE_LMO_GWP_HULU.md). Workbook gabungan tujuh sheet ialah [Rekap_Pareto_LMO_Lengkap.xlsx](Rekap_Pareto_LMO_Lengkap.xlsx); [Rekap_LCI_LMO.xlsx](Rekap_LCI_LMO.xlsx) dan [Rekap_PROC_LMO.xlsx](Rekap_PROC_LMO.xlsx) adalah hasil terpisah. Perbaikan LUC HULU sebelumnya didokumentasikan di [docs/KNOWLEDGE_LMO_LUC_HULU.md](docs/KNOWLEDGE_LMO_LUC_HULU.md). Generator yang sama menghasilkan ketiganya.

Untuk BMO, baca [docs/KNOWLEDGE_BMO.md](docs/KNOWLEDGE_BMO.md) sebelum menafsirkan hasil. Unggahan SimaPro BMO terbaru telah memperbaiki enam kategori HULU yang sebelumnya salah jenis; seluruh tujuh tabel kini lengkap. Hasilnya ada di [Rekap_Pareto_BMO.xlsx](Rekap_Pareto_BMO.xlsx) dan generatornya di [scripts/generate_bmo_pareto.py](scripts/generate_bmo_pareto.py).

Untuk SMO, baca [docs/KNOWLEDGE_SMO.md](docs/KNOWLEDGE_SMO.md) sebelum menafsirkan hasil. [Rekap_LCI_SMO.xlsx](Rekap_LCI_SMO.xlsx) memuat HULU, INTI, UT, serta penanda bahwa sumber `LCI DWN SMO` tidak tersedia; [Rekap_PROC_SMO.xlsx](Rekap_PROC_SMO.xlsx) memuat HULU, INTI, UT, DWN. Hasil gabungan dan jejak angka ada di [Rekap_Pareto_SMO_Lengkap.xlsx](Rekap_Pareto_SMO_Lengkap.xlsx). Generator: [scripts/generate_smo_pareto.py](scripts/generate_smo_pareto.py).

Untuk GMO, baca [docs/KNOWLEDGE_GMO.md](docs/KNOWLEDGE_GMO.md) sebelum menafsirkan hasil. [Rekap_LCI_GMO.xlsx](Rekap_LCI_GMO.xlsx) memuat HULU, INTI, UT; [Rekap_PROC_GMO.xlsx](Rekap_PROC_GMO.xlsx) memuat HULU, INTI, UT, DWN. LCI DWN sengaja tidak dibuat sesuai permintaan. Hasil gabungan dan jejak angka ada di [Rekap_Pareto_GMO_Lengkap.xlsx](Rekap_Pareto_GMO_Lengkap.xlsx). Generator: [scripts/generate_gmo_pareto.py](scripts/generate_gmo_pareto.py).
