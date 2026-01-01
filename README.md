# Excel-Access-ETL (VBA + Access + Power Query)

A portable Excel ↔ Access ETL demo using:
- **Excel VBA (ADODB)** to export/import
- **Access (.accdb)** as the storage layer
- **Power Query** refresh (best-effort) before export

This repo is **code-only**: it publishes the exported VBA modules (`.bas`) and documentation.
Excel workbooks (`.xlsm`) and Access databases (`.accdb`) are intentionally excluded via `.gitignore`.

## What it does
1. **ExportToAccess**
   - (Optional) refreshes Power Query queries (`pRegion`, `SalesData`) if they exist
   - exports rows from Excel table `SalesData` into Access table `tbl_Sales`
   - uses **transaction + rollback** for safety
2. **ImportFromAccess**
   - imports `tbl_Sales` into a new worksheet `Imported_Results`
   - uses `CopyFromRecordset` for fast transfers

## Expected schema

### Excel
- Worksheet: `Sheet1`
- Excel Table (ListObject): `SalesData`
- Columns: `ID`, `Product`, `Sales`, `Region`

### Access
- Database file: `ProjectDB.accdb`
- Table: `tbl_Sales`
- Fields: `ID`, `Product`, `Sales`, `Region`

## Database path resolution (portable, no hardcoded paths)
The macros locate the Access database using `ResolveAccessDbPath()` in this order:
1. Environment variable `ACCESS_DB_PATH` (full path to `ProjectDB.accdb`)
2. Same folder as the workbook
3. Common repo folders next to the workbook: `data\`, `db\`, `assets\`, `sample\`

## Requirements
- Excel desktop (VBA enabled)
- Access Database Engine / ACE OLEDB provider available
- VBA reference: **Microsoft ActiveX Data Objects** (6.1 or 2.8)

## How to use (local)
1. Create/open your `.xlsm`
2. Ensure your Access DB exists as `ProjectDB.accdb` (or set `ACCESS_DB_PATH`)
3. Import the modules from `src/vba` into VBA Editor:
   - `modETL_Helpers.bas`
   - `modExportToAccess.bas`
   - `modImportFromAccess.bas`
4. Run:
   - `ExportToAccess`
   - `ImportFromAccess`

## Source code
See: `src/vba/`
- `modETL_Helpers.bas`
- `modExportToAccess.bas`
- `modImportFromAccess.bas`

## Known limitations
- Export uses **DELETE + INSERT** (full refresh) rather than UPSERT
- Row-by-row INSERT is fine for small datasets; optimize for large volumes if needed
- Power Query refresh behavior varies by Excel version/build

## License
MIT (see `LICENSE`)

---

## 🚀 Task 16: Validation & Logging System (RECOMMENDED)

### Enhanced Module: modExportWithValidation.bas

**Production-grade ETL with comprehensive validation and audit trail.**

#### Key Features
- ✅ **Pre-flight Validation** - Validates all data before database operations
- ✅ **Transaction Safety** - ACID-compliant with automatic rollback on error
- ✅ **Intelligent UPSERT** - Updates existing records, inserts only new ones  
- ✅ **Dual Logging** - Audit trail in both Excel worksheet and Access database
- ✅ **Portable Paths** - Database resolves to workbook location automatically
- ✅ **SQL Injection Protection** - Proper parameter escaping
- ✅ **Progress Tracking** - 10-step status bar with detailed feedback

#### Required Access Table

Create this table in your Access database:

```sql
CREATE TABLE tbl_ETL_Log (
  LogID AUTOINCREMENT PRIMARY KEY,
  RunTimestamp DATETIME NOT NULL,
  Operation TEXT(50) NOT NULL,
  RecordsProcessed LONG,
  RecordsInserted LONG,
  RecordsUpdated LONG,
  RecordsFailed LONG,
  Status TEXT(20) NOT NULL,
  ErrorText MEMO,
  DurationSeconds DOUBLE
);

