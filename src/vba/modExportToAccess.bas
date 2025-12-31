Option Explicit
' Requires: modETL_Helpers (ResolveAccessDbPath, SafeValue, GetExcelBuildInfo)

Public Sub ExportToAccess()
'=================================================
' ETL EXPORT (Excel -> Access)
' - Refresh PQ parameter + data query
' - Transaction-safe export (DELETE + INSERT)
'
' Expected Excel Table: SalesData (ID, Product, Sales, Region)
' Expected Access Table: tbl_Sales (ID, Product, Sales, Region)
'
' DB resolution (no hardcoded paths):
' - ENV: ACCESS_DB_PATH
' - Same folder as workbook
' - \data \db \assets \sample beside workbook
'=================================================

    Dim conn As ADODB.Connection
    Dim strConn As String, strSQL As String, dbPath As String
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long, recordCount As Long
    Dim startTime As Double

    On Error GoTo ErrorHandler

    startTime = Timer
    recordCount = 0

    '--- Resolve DB path (portable) ---
    dbPath = ResolveAccessDbPath("ProjectDB.accdb")
    If Len(dbPath) = 0 Then
        MsgBox "Access database not found." & vbCrLf & vbCrLf & _
               "Fix one of these:" & vbCrLf & _
               "1) Put ProjectDB.accdb next to the workbook (or in \data \db \assets \sample)." & vbCrLf & _
               "2) Set ENV var ACCESS_DB_PATH to the full .accdb path." & vbCrLf & vbCrLf & _
               "Build: " & GetExcelBuildInfo(), _
               vbCritical, "Missing Database"
        Exit Sub
    End If

    '--- Excel objects ---
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set tbl = ws.ListObjects("SalesData")

    '=== Refresh Power Query (best effort) ===
    Application.StatusBar = "Refreshing Power Query..."
    On Error Resume Next

    'Parameter query
    ThisWorkbook.Queries("pRegion").Refresh
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Connections("Query - pRegion").Refresh
    End If

    DoEvents

    'Main query (if you have it). If you don't, harmless.
    ThisWorkbook.Queries("SalesData").Refresh
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Connections("Query - SalesData").Refresh
    End If

    On Error GoTo ErrorHandler

    '=== Connect to Access ===
    Application.StatusBar = "Connecting to Access..."
    Set conn = New ADODB.Connection
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    conn.Open strConn

    '=== Transaction ===
    conn.BeginTrans
    Application.StatusBar = "Exporting..."

    'Clean slate
    conn.Execute "DELETE FROM tbl_Sales"

    'Insert rows
    For i = 1 To tbl.ListRows.Count

        'Excel columns: 1=ID, 2=Product, 3=Sales, 4=Region
        strSQL = "INSERT INTO tbl_Sales (ID, Product, Sales, Region) VALUES (" & _
                 tbl.DataBodyRange(i, 1).Value & ", " & _
                 SafeValue(tbl.DataBodyRange(i, 2)) & ", " & _
                 tbl.DataBodyRange(i, 3).Value & ", " & _
                 SafeValue(tbl.DataBodyRange(i, 4)) & ")"

        conn.Execute strSQL
        recordCount = recordCount + 1

        If i Mod 200 = 0 Then
            Application.StatusBar = "Exporting... " & i & " / " & tbl.ListRows.Count
        End If
    Next i

    conn.CommitTrans

CleanExit:
    On Error Resume Next
    Application.StatusBar = False

    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing

    If Err.Number = 0 Then
        MsgBox "Export complete." & vbCrLf & _
               "Records: " & recordCount & vbCrLf & _
               "Time: " & Round(Timer - startTime, 2) & " sec" & vbCrLf & _
               "DB: " & dbPath & vbCrLf & _
               "Build: " & GetExcelBuildInfo(), _
               vbInformation, "Export Complete"
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.RollbackTrans
    End If

    MsgBox "Export failed at record " & i & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "DB: " & dbPath & vbCrLf & _
           "Build: " & GetExcelBuildInfo(), _
           vbCritical, "Export Error"

    Resume CleanExit
End Sub


