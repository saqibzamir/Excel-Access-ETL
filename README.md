\# Excel-Access-ETL



A lightweight Excel VBA + Power Query + Access ETL demo:



\- Excel Table (`SalesData`) is filtered via Power Query (pRegion driven by Sheet1!F2)

\- VBA exports filtered rows to Access table (`tbl\_Sales`) using ADODB (transaction + error handling)

\- VBA imports from Access back into Excel (`Imported\_Results`) using `CopyFromRecordset`



\## Requirements

\- Microsoft Excel (tested: Excel 16.0 Build 19426)

\- Microsoft Access Database Engine (ACE OLEDB 12.0)

\- VBA reference enabled: \*\*Microsoft ActiveX Data Objects 6.1 Library\*\* (or similar)



\## Files / Structure

\- `src/vba/` contains exported VBA modules (.bas)

\- `ProjectDB.accdb` is expected to be in the same folder as the workbook (if you use relative paths)



\## How to Run

1\. Open the workbook

2\. Enter a region value in \*\*Sheet1!F2\*\* (e.g., North / South)

3\. Run `ExportToAccess` (Alt+F8)

4\. Run `ImportFromAccess` (Alt+F8)



\## Known Limitations

\- Uses DELETE-then-INSERT (full refresh approach)

\- Row-by-row INSERT (fine for small datasets; optimize for large volumes)

\- Power Query refresh behavior can vary by Excel version



