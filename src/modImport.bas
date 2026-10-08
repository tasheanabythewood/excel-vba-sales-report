Attribute VB_Name = "modImport"
Option Explicit

' Imports a CSV into tblTransactions. Returns True only if the import completed.
Public Function ImportTransactions() As Boolean
    Dim filePath As String
    Dim src As Workbook
    Dim srcData As Range
    Dim tbl As ListObject
    Dim nRows As Long
    Dim problems As Long
    Dim t As Double

    t = Timer

    ' 1. Ask the user for a file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a transactions CSV"
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function         ' user clicked Cancel (returns False)
        filePath = .SelectedItems(1)
    End With

    On Error GoTo ImportFailed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 2. Open the file read-only and measure the data
    Set src = Workbooks.Open(fileName:=filePath, ReadOnly:=True)
    Set srcData = src.Worksheets(1).Range("A1").CurrentRegion
    nRows = srcData.Rows.Count - 1

    ' 3. Validate BEFORE touching the existing data
    problems = ValidateTransactions(srcData, Dir(filePath))
    If problems > 0 Then
        src.Close SaveChanges:=False
        Set src = Nothing
        WriteLog Dir(filePath), nRows, "REJECTED (" & problems & " problems)"
        Application.ScreenUpdating = True
        ThisWorkbook.Worksheets("Exceptions").ListObjects("tblExceptions").Range.Columns.AutoFit
        ThisWorkbook.Worksheets("Exceptions").Activate          ' show the user what's wrong
        MsgBox "This file was NOT imported: " & problems & " problem(s) found." & vbNewLine & _
               "Your existing data has not been changed." & vbNewLine & vbNewLine & _
               "See the Exceptions sheet for details.", vbExclamation, "Import rejected"
        GoTo CleanExit
    End If

    ' 4. Empty the table (replace, don't append)
    Set tbl = ThisWorkbook.Worksheets("Transactions").ListObjects("tblTransactions")
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete

    ' 5. Copy the 8 source columns in under the header, then resize the table
    tbl.HeaderRowRange.Offset(1).Resize(nRows, 8).Value = _
        srcData.Offset(1).Resize(nRows, 8).Value
    tbl.Resize tbl.Range.Resize(nRows + 1, tbl.ListColumns.Count)

    ' 6. Close the source file without saving
    src.Close SaveChanges:=False
    Set src = Nothing

    ' 7. Log it, recalculate once, then point the report at the imported month
    WriteLog Dir(filePath), nRows, "OK"
    Application.Calculation = xlCalculationAutomatic
    RefreshReport GetDataMonth()

    Debug.Print "Import took " & Format(Timer - t, "0.0") & " seconds"
    ImportTransactions = True                    ' tell the caller it worked
    MsgBox nRows & " rows imported from " & Dir(filePath), vbInformation, "Import complete"

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function

ImportFailed:
    If Not src Is Nothing Then src.Close SaveChanges:=False
    MsgBox "The import failed and was not completed." & vbNewLine & vbNewLine & _
           "Details: " & Err.Description, vbCritical, "Import error"
    Resume CleanExit
End Function

' Button-friendly version (buttons need a Sub)
Public Sub ImportData()
    ImportTransactions
End Sub

' Adds one row to the Import Log
Private Sub WriteLog(ByVal fileName As String, ByVal rowCount As Long, ByVal status As String)
    Dim logRow As ListRow
    Set logRow = ThisWorkbook.Worksheets("Import Log").ListObjects("tblImportLog").ListRows.Add
    logRow.Range(1, 1).Value = Now
    logRow.Range(1, 2).Value = fileName
    logRow.Range(1, 3).Value = rowCount
    logRow.Range(1, 4).Value = status
End Sub

