# Konteks proyek

Untuk pekerjaan terkait rekap LMO, baca [docs/KNOWLEDGE_LMO.md](docs/KNOWLEDGE_LMO.md) sebelum mengubah analisis atau menafsirkan workbook. Dokumen tersebut mencatat sumber, aturan Pareto, pemetaan kategori, duplikasi data, dan batas atribusi LCI. Hasil rinci ada di [Rekap_Pareto_LMO.xlsx](Rekap_Pareto_LMO.xlsx).

Untuk proses hilir, baca juga [docs/KNOWLEDGE_PROC_DWN.md](docs/KNOWLEDGE_PROC_DWN.md). Sheet SimaPro bernama `PROC DWN`; hasilnya ada di [Hotspot_LMO_PROC_DWN.xlsx](Hotspot_LMO_PROC_DWN.xlsx), dan generatornya di [scripts/generate_proc_dwn.py](scripts/generate_proc_dwn.py).

Untuk rekap ulang dari unggahan SimaPro terbaru, baca [docs/KNOWLEDGE_LMO_5_SHEET.md](docs/KNOWLEDGE_LMO_5_SHEET.md). Hasil lima sheet INTI, UT, DOWN berada di [Rekap_Pareto_LMO_5_Sheet.xlsx](Rekap_Pareto_LMO_5_Sheet.xlsx); tersedia juga [Rekap_LCI_LMO.xlsx](Rekap_LCI_LMO.xlsx) dan [Rekap_PROC_LMO.xlsx](Rekap_PROC_LMO.xlsx). Generator ada di [scripts/generate_lmo_five_sheets.py](scripts/generate_lmo_five_sheets.py). Sheet sumber hilir bernama `PROC DWN LMO`, sedangkan sheet hasil `PROC DOWN LMO`. Unggahan terbaru telah menambahkan WSF pada `PROC INTI LMO`.

Untuk BMO, baca [docs/KNOWLEDGE_BMO.md](docs/KNOWLEDGE_BMO.md) sebelum menafsirkan hasil. Unggahan SimaPro BMO terbaru telah memperbaiki enam kategori HULU yang sebelumnya salah jenis; seluruh tujuh tabel kini lengkap. Hasilnya ada di [Rekap_Pareto_BMO.xlsx](Rekap_Pareto_BMO.xlsx) dan generatornya di [scripts/generate_bmo_pareto.py](scripts/generate_bmo_pareto.py).
