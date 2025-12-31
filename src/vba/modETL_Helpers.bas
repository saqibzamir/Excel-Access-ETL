Option Explicit
' ETL Helper Functions:
' - ResolveAccessDbPath: portable DB resolution (no hardcoded local paths)
' - EscapeSQL / SafeValue: SQL safety
' - GetExcelBuildInfo: environment logging

Public Function ResolveAccessDbPath(Optional ByVal dbFileName As String = "ProjectDB.accdb") As String
    'Resolution order:
    '1) ENV var override: ACCESS_DB_PATH (full path to .accdb)
    '2) Same folder as workbook
    '3) Common repo folders: \data \db \assets \sample
    'Returns "" if not found or workbook not saved.

    Dim base As String
    Dim p As String
    Dim candidates As Variant
    Dim i As Long

    '1) Environment variable override
    p = Trim$(Environ$("ACCESS_DB_PATH"))
    If Len(p) > 0 Then
        If LCase$(Right$(p, 6)) = ".accdb" Then
            If Dir$(p) <> "" Then
                ResolveAccessDbPath = p
                Exit Function
            End If
        End If
    End If

    'Workbook must be saved for Path to exist
    base = ThisWorkbook.Path
    If Len(base) = 0 Then
        ResolveAccessDbPath = ""
        Exit Function
    End If

    '2) Same folder as workbook
    p = base & Application.PathSeparator & dbFileName
    If Dir$(p) <> "" Then
        ResolveAccessDbPath = p
        Exit Function
    End If

    '3) Common repo subfolders
    candidates = Array( _
        base & Application.PathSeparator & "data" & Application.PathSeparator & dbFileName, _
        base & Application.PathSeparator & "db" & Application.PathSeparator & dbFileName, _
        base & Application.PathSeparator & "assets" & Application.PathSeparator & dbFileName, _
        base & Application.PathSeparator & "sample" & Application.PathSeparator & dbFileName _
    )

    For i = LBound(candidates) To UBound(candidates)
        If Dir$(CStr(candidates(i))) <> "" Then
            ResolveAccessDbPath = CStr(candidates(i))
            Exit Function
        End If
    Next i

    ResolveAccessDbPath = ""
End Function

Public Function EscapeSQL(ByVal txt As Variant) As String
    'Escapes single quotes for SQL string literals
    'Mike's -> Mike''s
    If IsNull(txt) Or IsEmpty(txt) Then
        EscapeSQL = ""
    Else
        EscapeSQL = Replace(CStr(txt), "'", "''")
    End If
End Function

Public Function SafeValue(ByVal cell As Range) As String
    'Returns SQL-safe values for Access SQL:
    ' - NULL
    ' - Dates wrapped in #...#
    ' - Numbers as-is
    ' - Text wrapped in '...' with escaped quotes

    If IsEmpty(cell.Value) Or IsNull(cell.Value) Then
        SafeValue = "NULL"
    ElseIf IsDate(cell.Value) Then
        SafeValue = "#" & Format$(cell.Value, "yyyy-mm-dd") & "#"
    ElseIf IsNumeric(cell.Value) And Not IsDate(cell.Value) Then
        SafeValue = CStr(cell.Value)
    Else
        SafeValue = "'" & EscapeSQL(cell.Value) & "'"
    End If
End Function

Public Function GetExcelBuildInfo() As String
    GetExcelBuildInfo = "Excel " & Application.Version & " Build " & Application.Build
End Function


