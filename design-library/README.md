# Standard Design Library (Singapore)

This is a library of standard civil designs that are pre-checked against Singapore requirements. You use it to call up designs for tender drawings.

| File | What it is |
|---|---|
| `Design_Library_Register.xlsx` | The registers (drain, RC wall, embedded wall, manhole, ERSS pit, linkway, others), Design Selector, Tender Pack, CAD Export, Coverage, the SG criteria and the CAD standard |
| `TenderSheets.lsp` | AutoCAD routine (`TPSHEETS`). It reads the CAD_Export text file, copies the title-block layout for each sheet, inserts the detail DWGs at 1:n and writes the detail titles |
| `design-selector-prototype.html` | A clickable prototype of the dashboard to build in Power Apps. Its data is generated from the workbook |
| `tools/` | `build_register.py` rebuilds the workbook template. `gen_prototype.py` regenerates the prototype from the workbook |

Each design's file set is stored in `<library root>\<register folder>\<Design ID>\`:
`<ID>_CAD.dwg`, `<ID>_DWG.pdf`, `<ID>_STAAD.std`, `<ID>_CALC.xlsx` and `<ID>_CALC.pdf`.

Every row in the example data is illustrative. Before go-live, confirm every SG_Criteria row marked **VERIFY** against the current PUB, LTA and BCA documents.
