# Standard Design Library (Singapore)

This is a library of standard civil designs that are pre-checked against Singapore requirements. You use it to call up designs for tender drawings.

| File | What it is |
|---|---|
| `Design_Library_Register.xlsx` | **The database.** Registers (drain, RC wall, embedded wall, manhole, ERSS pit, linkway, others), Projects, Rebar_Schedule, Revision_Log, Data_Dictionary, plus Design Selector, Tender Pack, CAD Export, Coverage, SG criteria and CAD standard |
| `TenderSheets.lsp` | AutoCAD routine (`TPSHEETS`). It reads the CAD_Export text file, copies the title-block layout for each sheet, inserts the detail DWGs at 1:n and writes the detail titles |
| `design-library-search.html` | Search tool. Opens the Excel register in the browser (nothing is uploaded) and shows each design's full record. Ships with the example data |
| `powerapps/` | Power Apps version of the search tool: setup guide, load formula and pasteable screens (`.pa.yaml`) |
| `tools/` | `build_register.py` rebuilds the workbook template. `gen_prototype.py` rebuilds the search page with the workbook's data embedded |

Each design's file set is stored in `<library root>\<register folder>\<Design ID>\`:
`<ID>_CAD.dwg`, `<ID>_DWG.pdf`, `<ID>_STAAD.std`, `<ID>_CALC.xlsx` and `<ID>_CALC.pdf`.

Every row in the example data is illustrative. Before go-live, confirm every SG_Criteria row marked **VERIFY** against the current PUB, LTA and BCA documents.
