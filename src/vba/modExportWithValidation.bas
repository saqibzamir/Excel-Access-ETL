 ' ============================================================================
' PROFESSIONAL EXCEL-TO-ACCESS ETL SYSTEM
' ============================================================================
' Description: Robust data synchronization between Excel and Microsoft Access
'              with validation, transaction safety, and comprehensive logging
'
' Features:
'   - Pre-flight data validation with detailed error reporting
'   - ACID-compliant transactions (rollback on errors)
'   - Intelligent UPSERT (UPDATE existing + INSERT new records)
'   - Dual audit logging (Excel worksheet + Access table)
'   - Portable path resolution (no hardcoded paths)
'   - SQL injection protection
'   - Production-grade error handling
'
' Requirements:
'   - Microsoft Excel 2013+ with VBA enabled
'   - Microsoft Access 2013+ database
'   - Microsoft ActiveX Data Objects 2.8+ Library
'
' License: MIT (modify freely)
' ============================================================================

Option Explicit

' ============================================================================
' CONFIGURATION - CUSTOMIZE FOR YOUR ENVIRONMENT
' ============================================================================
Private Const DB_FILENAME As String = "SalesDatabase.accdb"     ' Access database filename
Private Const LOG_SHEET_NAME As String = "ETL_Log"              ' Excel log sheet name
Private Const SOURCE_SHEET_NAME As String = "Sheet 1"           ' Excel source data sheet
Private Const TARGET_TABLE As String = "tbl_Sales"              ' Access destination table
Private Const LOG_TABLE As String = "tbl_ETL_Log"               ' Access audit log table
Private Const STAGING_TABLE As String = "tbl_Staging"           ' Temporary staging table

' Field Mapping (Column positions in Excel)
Private Const COL_ID As Long = 1                                ' Column A: Record ID
Private Const COL_PRODUCT As Long = 2                           ' Column B: Product Name
Private Const COL_SALES As Long = 3                             ' Column C: Sales Amount
Private Const COL_REGION As Long = 4                            ' Column D: Region
Private Const COL_STATUS As Long = 5                            ' Column E: Export Status

' ============================================================================
' PUBLIC ENTRY POINT - Main Export Procedure
' ============================================================================
Public Sub ExportSalesData()
    Dim startTime As Double: startTime = Timer
    Dim conn As Object ' ADODB.Connection
    Dim dbPath As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim recordsProcessed As Long, recordsInserted As Long
    Dim recordsUpdated As Long, recordsFailed As Long
    Dim errorLog As String
    Dim validationErrors As Collection
    
    On Error GoTo ErrorHandler
    
    ' STEP 1: LOCATE DATABASE
    Application.StatusBar = "Step 1/10: Locating database..."
    dbPath = ResolveDatabasePath()
    
    If dbPath = "" Then
        MsgBox "Database Not Found" & vbCrLf & vbCrLf & _
               "Please place " & DB_FILENAME & " in the same folder as this Excel file," & vbCrLf & _
               "or select it when prompted.", _
               vbCritical, "Database Missing"
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' STEP 2: ENSURE LOG INFRASTRUCTURE EXISTS
    Application.StatusBar = "Step 2/10: Preparing log sheet..."
    If Not SheetExists(LOG_SHEET_NAME) Then CreateLogSheet
    
    ' STEP 3: VALIDATE SOURCE DATA
    Application.StatusBar = "Step 3/10: Validating data..."
    Set ws = ThisWorkbook.Sheets(SOURCE_SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "No Data Found" & vbCrLf & vbCrLf & _
               "Please add sales records to " & SOURCE_SHEET_NAME & " starting in row 2.", _
               vbExclamation, "No Data"
        Application.StatusBar = False
        Exit Sub
    End If
    
    Set validationErrors = New Collection
    
    If Not ValidateAllRows(ws, lastRow, validationErrors) Then
        recordsFailed = validationErrors.Count
        errorLog = BuildErrorReport(validationErrors)
        
        Call WriteExcelLog("Export Sales", 0, 0, 0, recordsFailed, _
                          "Validation Failed", Timer - startTime, errorLog)
        
        MsgBox "Data Validation Failed" & vbCrLf & vbCrLf & _
               "Errors Found: " & recordsFailed & vbCrLf & vbCrLf & _
               Left(errorLog, 400) & vbCrLf & vbCrLf & _
               "Check " & LOG_SHEET_NAME & " sheet for complete error details.", _
               vbExclamation, "Validation Failed"
        
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' STEP 4: ESTABLISH DATABASE CONNECTION
    Application.StatusBar = "Step 4/10: Connecting to database..."
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    ' STEP 5: BEGIN TRANSACTION (ensures all-or-nothing operation)
    Application.StatusBar = "Step 5/10: Starting transaction..."
    conn.BeginTrans
    
    ' STEP 6: CREATE STAGING TABLE
    Application.StatusBar = "Step 6/10: Creating staging area..."
    Call CreateStagingTable(conn)
    
    ' STEP 7: EXPORT TO STAGING
    Application.StatusBar = "Step 7/10: Exporting data..."
    Call ExportToStaging(conn, ws, lastRow, recordsProcessed, recordsFailed, errorLog)
    
    ' STEP 8: PERFORM UPSERT
    Application.StatusBar = "Step 8/10: Updating database..."
    Call PerformUpsert(conn, recordsInserted, recordsUpdated)
    
    ' STEP 9: COMMIT TRANSACTION
    Application.StatusBar = "Step 9/10: Committing changes..."
    conn.CommitTrans
    
    ' STEP 10: CLEANUP
    Application.StatusBar = "Step 10/10: Cleanup..."
    On Error Resume Next
    conn.Execute "DROP TABLE " & STAGING_TABLE
    On Error GoTo ErrorHandler
    
    ' LOG SUCCESS AND NOTIFY USER
    Dim duration As Double: duration = Timer - startTime
    
    Call WriteExcelLog("Export Sales", recordsProcessed, recordsInserted, _
                      recordsUpdated, recordsFailed, "Success", duration, errorLog)
    Call WriteAccessLog(conn, "Export Sales", recordsProcessed, recordsInserted, _
                       recordsUpdated, recordsFailed, "Success", errorLog, duration)
    
    conn.Close
    Set conn = Nothing
    Application.StatusBar = False
    
    ' Success Summary
    Dim msg As String
    msg = "EXPORT COMPLETED SUCCESSFULLY!" & vbCrLf & vbCrLf & _
          "SUMMARY:" & vbCrLf & _
          "-------------------" & vbCrLf & _
          "Processed: " & recordsProcessed & " records" & vbCrLf & _
          "Inserted: " & recordsInserted & " new" & vbCrLf & _
          "Updated: " & recordsUpdated & " existing" & vbCrLf & _
          "Failed: " & recordsFailed & vbCrLf & _
          "Duration: " & Format(duration, "0.00") & " seconds" & vbCrLf & vbCrLf & _
          "Check Access " & TARGET_TABLE & " to verify changes."
    
    If recordsFailed > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Warning: Some records failed - see " & LOG_SHEET_NAME
        MsgBox msg, vbExclamation, "Export Completed with Warnings"
    Else
        MsgBox msg, vbInformation, "Export Successful"
    End If
    
    Exit Sub
    
' ============================================================================
' ERROR HANDLER
' ============================================================================
ErrorHandler:
    Application.StatusBar = False
    
    If Not conn Is Nothing Then
        On Error Resume Next
        If conn.State = 1 Then
            conn.RollbackTrans
            conn.Close
        End If
        On Error GoTo 0
    End If
    
    Dim criticalError As String
    criticalError = "CRITICAL ERROR: " & Err.Description & " (Error #" & Err.Number & ")"
    errorLog = errorLog & vbCrLf & criticalError
    
    Call WriteExcelLog("Export Sales", recordsProcessed, recordsInserted, _
                      recordsUpdated, recordsFailed, "Failed", _
                      Timer - startTime, errorLog)
    
    MsgBox "EXPORT FAILED" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Transaction rolled back - no changes were made to the database." & vbCrLf & _
           "Check " & LOG_SHEET_NAME & " for detailed error information.", _
           vbCritical, "Export Failed"
End Sub

' ============================================================================
' VALIDATION FUNCTIONS
' ============================================================================
Private Function ValidateAllRows(ws As Worksheet, lastRow As Long, _
                                errors As Collection) As Boolean
    Dim i As Long
    Dim id As Variant, product As Variant, sales As Variant, region As Variant
    Dim isValid As Boolean: isValid = True
    
    For i = 2 To lastRow
        id = ws.Cells(i, COL_ID).Value
        product = ws.Cells(i, COL_PRODUCT).Value
        sales = ws.Cells(i, COL_SALES).Value
        region = ws.Cells(i, COL_REGION).Value
        
        ' RULE 1: ID Validation
        If IsEmpty(id) Then
            errors.Add "Row " & i & ": ID is empty"
            isValid = False
        ElseIf Not IsNumeric(id) Then
            errors.Add "Row " & i & ": ID must be numeric"
            isValid = False
        ElseIf CLng(id) <= 0 Then
            errors.Add "Row " & i & ": ID must be positive"
            isValid = False
        End If
        
        ' RULE 2: Product Validation
        If IsEmpty(product) Or Trim(CStr(product)) = "" Then
            errors.Add "Row " & i & ": Product name is empty"
            isValid = False
        ElseIf Len(Trim(CStr(product))) < 2 Then
            errors.Add "Row " & i & ": Product name too short (min 2 characters)"
            isValid = False
        End If
        
        ' RULE 3: Sales Amount Validation
        If IsEmpty(sales) Then
            errors.Add "Row " & i & ": Sales amount is empty"
            isValid = False
        ElseIf Not IsNumeric(sales) Then
            errors.Add "Row " & i & ": Sales must be numeric"
            isValid = False
        ElseIf CDbl(sales) < 0 Then
            errors.Add "Row " & i & ": Sales cannot be negative"
            isValid = False
        ElseIf CDbl(sales) > 1000000 Then
            errors.Add "Row " & i & ": Sales exceeds $1M limit"
            isValid = False
        End If
        
        ' RULE 4: Region Validation
        If IsEmpty(region) Or Trim(CStr(region)) = "" Then
            errors.Add "Row " & i & ": Region is empty"
            isValid = False
        End If
        
        ' Mark validation status in Excel
        If IsEmpty(id) Or IsEmpty(product) Or IsEmpty(sales) Or IsEmpty(region) Then
            ws.Cells(i, COL_STATUS).Value = "Invalid"
            ws.Cells(i, COL_STATUS).Interior.Color = RGB(255, 200, 200)
        Else
            ws.Cells(i, COL_STATUS).Value = "Valid"
            ws.Cells(i, COL_STATUS).Interior.Color = RGB(200, 255, 200)
        End If
    Next i
    
    ValidateAllRows = isValid
End Function

Private Function BuildErrorReport(errors As Collection) As String
    Dim msg As String, i As Long
    
    If errors.Count = 0 Then
        BuildErrorReport = ""
        Exit Function
    End If
    
    msg = "Validation Errors (" & errors.Count & "):" & vbCrLf
    For i = 1 To WorksheetFunction.Min(errors.Count, 20)
        msg = msg & "- " & errors(i) & vbCrLf
    Next i
    
    If errors.Count > 20 Then
        msg = msg & "... and " & (errors.Count - 20) & " more errors"
    End If
    
    BuildErrorReport = msg
End Function

' ============================================================================
' DATABASE OPERATIONS
' ============================================================================
Private Sub CreateStagingTable(conn As Object)
    On Error Resume Next
    conn.Execute "DROP TABLE " & STAGING_TABLE
    On Error GoTo 0
    
    Dim sql As String
    sql = "CREATE TABLE " & STAGING_TABLE & " (" & _
          "ID LONG PRIMARY KEY, " & _
          "Product TEXT(100), " & _
          "Sales DOUBLE, " & _
          "Region TEXT(50)" & _
          ")"
    
    conn.Execute sql
End Sub

Private Sub ExportToStaging(conn As Object, ws As Worksheet, lastRow As Long, _
                           ByRef processed As Long, ByRef failed As Long, _
                           ByRef errorLog As String)
    Dim i As Long
    Dim sql As String
    Dim id As Long, product As String, sales As Double, region As String
    
    processed = 0
    failed = 0
    
    For i = 2 To lastRow
        On Error Resume Next
        
        id = CLng(ws.Cells(i, COL_ID).Value)
        product = Trim(CStr(ws.Cells(i, COL_PRODUCT).Value))
        sales = CDbl(ws.Cells(i, COL_SALES).Value)
        region = Trim(CStr(ws.Cells(i, COL_REGION).Value))
        
        If Err.Number = 0 Then
            ' SQL injection protection
            product = Replace(product, "'", "''")
            region = Replace(region, "'", "''")
            
            sql = "INSERT INTO " & STAGING_TABLE & " (ID, Product, Sales, Region) " & _
                  "VALUES (" & id & ", '" & product & "', " & sales & ", '" & region & "')"
            
            conn.Execute sql
            
            If Err.Number = 0 Then
                processed = processed + 1
                ws.Cells(i, COL_STATUS).Value = "Exported"
                ws.Cells(i, COL_STATUS).Interior.Color = RGB(200, 255, 200)
            Else
                failed = failed + 1
                errorLog = errorLog & "Row " & i & ": " & Err.Description & " | "
                ws.Cells(i, COL_STATUS).Value = "Failed"
                ws.Cells(i, COL_STATUS).Interior.Color = RGB(255, 200, 200)
                Err.Clear
            End If
        Else
            failed = failed + 1
            errorLog = errorLog & "Row " & i & ": Conversion error | "
            ws.Cells(i, COL_STATUS).Value = "Invalid"
            ws.Cells(i, COL_STATUS).Interior.Color = RGB(255, 200, 200)
            Err.Clear
        End If
        
        On Error GoTo 0
    Next i
End Sub

Private Sub PerformUpsert(conn As Object, ByRef inserted As Long, ByRef updated As Long)
    Dim sql As String
    
    ' UPDATE existing records
    sql = "UPDATE " & TARGET_TABLE & " INNER JOIN " & STAGING_TABLE & " " & _
          "ON " & TARGET_TABLE & ".ID = " & STAGING_TABLE & ".ID " & _
          "SET " & TARGET_TABLE & ".Product = " & STAGING_TABLE & ".Product, " & _
          TARGET_TABLE & ".Sales = " & STAGING_TABLE & ".Sales, " & _
          TARGET_TABLE & ".Region = " & STAGING_TABLE & ".Region"
    
    conn.Execute sql, updated
    
    ' INSERT new records
    sql = "INSERT INTO " & TARGET_TABLE & " (ID, Product, Sales, Region) " & _
          "SELECT " & STAGING_TABLE & ".ID, " & STAGING_TABLE & ".Product, " & _
          STAGING_TABLE & ".Sales, " & STAGING_TABLE & ".Region " & _
          "FROM " & STAGING_TABLE & " LEFT JOIN " & TARGET_TABLE & " " & _
          "ON " & STAGING_TABLE & ".ID = " & TARGET_TABLE & ".ID " & _
          "WHERE " & TARGET_TABLE & ".ID IS NULL"
    
    conn.Execute sql, inserted
End Sub

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================
Private Function ResolveDatabasePath() As String
    Dim excelFolder As String
    Dim testPath As String
    
    ' Strategy 1: Same folder as Excel workbook
    excelFolder = ThisWorkbook.Path
    If Right(excelFolder, 1) <> "\" Then excelFolder = excelFolder & "\"
    testPath = excelFolder & DB_FILENAME
    
    If Dir(testPath) <> "" Then
        ResolveDatabasePath = testPath
        Exit Function
    End If
    
    ' Strategy 2: Prompt user to locate file
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Locate " & DB_FILENAME
        .Filters.Clear
        .Filters.Add "Access Database", "*.accdb;*.mdb"
        .AllowMultiSelect = False
        .InitialFileName = excelFolder
        
        If .Show = -1 Then
            ResolveDatabasePath = .SelectedItems(1)
        Else
            ResolveDatabasePath = ""
        End If
    End With
End Function

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub CreateLogSheet()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = LOG_SHEET_NAME
    
    With ws
        .Cells(1, 1).Value = "Timestamp"
        .Cells(1, 2).Value = "Operation"
        .Cells(1, 3).Value = "Processed"
        .Cells(1, 4).Value = "Inserted"
        .Cells(1, 5).Value = "Updated"
        .Cells(1, 6).Value = "Failed"
        .Cells(1, 7).Value = "Status"
        .Cells(1, 8).Value = "Duration"
        .Cells(1, 9).Value = "Errors"
        
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:I1").Font.Color = RGB(255, 255, 255)
        .Columns("A:I").AutoFit
    End With
End Sub

' ============================================================================
' LOGGING FUNCTIONS
' ============================================================================
Private Sub WriteExcelLog(operation As String, processed As Long, inserted As Long, _
                         updated As Long, failed As Long, status As String, _
                         duration As Double, errors As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = operation
    ws.Cells(nextRow, 3).Value = processed
    ws.Cells(nextRow, 4).Value = inserted
    ws.Cells(nextRow, 5).Value = updated
    ws.Cells(nextRow, 6).Value = failed
    ws.Cells(nextRow, 7).Value = status
    ws.Cells(nextRow, 8).Value = Round(duration, 2)
    ws.Cells(nextRow, 9).Value = Left(errors, 32000)
    
    Select Case UCase(status)
        Case "SUCCESS": ws.Cells(nextRow, 7).Interior.Color = RGB(200, 255, 200)
        Case "FAILED": ws.Cells(nextRow, 7).Interior.Color = RGB(255, 200, 200)
        Case "VALIDATION FAILED": ws.Cells(nextRow, 7).Interior.Color = RGB(255, 220, 150)
        Case Else: ws.Cells(nextRow, 7).Interior.Color = RGB(220, 220, 220)
    End Select
    
    ws.Columns("A:I").AutoFit
End Sub

Private Sub WriteAccessLog(conn As Object, operation As String, processed As Long, _
                          inserted As Long, updated As Long, failed As Long, _
                          status As String, errors As String, duration As Double)
    Dim sql As String
    
    errors = Replace(Left(errors, 5000), "'", "''")
    operation = Replace(operation, "'", "''")
    
    sql = "INSERT INTO " & LOG_TABLE & " (RunTimestamp, Operation, RecordsProcessed, " & _
          "RecordsInserted, RecordsUpdated, RecordsFailed, Status, ErrorText, DurationSeconds) " & _
          "VALUES (Now(), '" & operation & "', " & processed & ", " & inserted & ", " & _
          updated & ", " & failed & ", '" & status & "', '" & errors & "', " & duration & ")"
    
    On Error Resume Next
    conn.Execute sql
    On Error GoTo 0
End Sub

' ============================================================================
' OPTIONAL DIAGNOSTIC TOOL
' ============================================================================
Public Sub ShowSystemConfiguration()
    Dim dbPath As String
    dbPath = ResolveDatabasePath()
    
    Dim info As String
    info = "SYSTEM CONFIGURATION" & vbCrLf & _
           String(50, "=") & vbCrLf & vbCrLf & _
           "Excel Workbook:" & vbCrLf & _
           ThisWorkbook.FullName & vbCrLf & vbCrLf & _
           "Access Database:" & vbCrLf & _
           IIf(dbPath <> "", dbPath, "NOT FOUND") & vbCrLf & vbCrLf & _
           "Status Checks:" & vbCrLf & _
           "- Database Located: " & IIf(dbPath <> "", "YES", "NO") & vbCrLf & _
           "- Log Sheet Exists: " & IIf(SheetExists(LOG_SHEET_NAME), "YES", "NO") & vbCrLf & _
           "- Source Sheet Exists: " & IIf(SheetExists(SOURCE_SHEET_NAME), "YES", "NO") & vbCrLf & vbCrLf & _
           "Configuration:" & vbCrLf & _
           "- DB Filename: " & DB_FILENAME & vbCrLf & _
           "- Target Table: " & TARGET_TABLE & vbCrLf & _
           "- Log Table: " & LOG_TABLE
    
    MsgBox info, vbInformation, "System Configuration"
End Sub


