Option Explicit
' Requires: modETL_Helpers (ResolveAccessDbPath, GetExcelBuildInfo)

Public Sub ImportFromAccess()
'=================================================
' ETL IMPORT (Access -> Excel)
' - Pulls tbl_Sales into fresh sheet Imported_Results
' - Uses CopyFromRecordset for speed
'
' DB resolution (no hardcoded paths):
' - ENV: ACCESS_DB_PATH
' - Same folder as workbook
' - \data \db \assets \sample beside workbook
'=================================================

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim strConn As String, dbPath As String
    Dim i As Long, startTime As Double

    On Error GoTo ErrorHandler

    startTime = Timer

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

    '--- Replace previous output (safe alerts handling) ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Imported_Results").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    'Create new output sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Imported_Results"

    'Connect
    Set conn = New ADODB.Connection
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    conn.Open strConn

    'Query
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tbl_Sales", conn, adOpenStatic, adLockReadOnly

    'Headers
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i

    'Data
    If Not rs.EOF Then
        ws.Range("A2").CopyFromRecordset rs
    End If

    'Format
    ws.Columns.AutoFit
    ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = "tbl_Imported"

CleanExit:
    On Error Resume Next
    Application.DisplayAlerts = True

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set rs = Nothing
    Set conn = Nothing

    If Err.Number = 0 Then
        MsgBox "Import complete." & vbCrLf & _
               "Time: " & Round(Timer - startTime, 2) & " sec" & vbCrLf & _
               "DB: " & dbPath & vbCrLf & _
               "Build: " & GetExcelBuildInfo(), _
               vbInformation, "Import Complete"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Import failed." & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "DB: " & dbPath & vbCrLf & _
           "Build: " & GetExcelBuildInfo(), _
           vbCritical, "Import Error"
    Resume CleanExit
End Sub


