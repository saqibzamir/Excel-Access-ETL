\# Excel-to-Access ETL System

---

## 🎯 Task 16: Production Dashboard & Path Persistence

### Smart Database Path Resolution

**Problem Solved:** No more repeated database prompts!

**How It Works:**
1. **First check**: Same folder as workbook → Auto-detect
2. **Second check**: Previously saved path (stored in K1 cell)
3. **Third option**: User prompt → Save selection for future

**Path Persistence:**
```vba
' First export: User prompted to locate ProjectDB.accdb
' Selection saved in ETL_Log!K1 + named range ETL_DB_PATH
' Future exports: Uses saved path (no prompting)
' Database moved? Run ClearSavedDatabasePath()


Professional data synchronization tool for Excel and Microsoft Access with validation, logging, and transaction safety.



\## Features



\- ✅ Pre-flight data validation

\- ✅ ACID-compliant transactions

\- ✅ Intelligent UPSERT (update + insert)

\- ✅ Dual audit logging (Excel + Access)

\- ✅ Portable (no hardcoded paths)

\- ✅ SQL injection protection

\- ✅ Comprehensive error handling



\## Quick Start



\### Prerequisites

\- Microsoft Excel 2013+

\- Microsoft Access 2013+

\- Enable macros in Excel



\### Setup

1\. Place `SalesDatabase.accdb` in same folder as Excel workbook

2\. Add VBA code to new module in Excel

3\. Ensure source data sheet is named "Sheet 1" with columns:

&nbsp;  - Column A: ID (numeric)

&nbsp;  - Column B: Product (text)

&nbsp;  - Column C: Sales (numeric)

&nbsp;  - Column D: Region (text)



\### Run Export

```vba

' From VBA Immediate Window (Ctrl+G)

ExportSalesData



