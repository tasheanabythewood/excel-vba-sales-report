Attribute VB_Name = "modValidate"
Option Explicit

Private Const MAX_EXCEPTIONS As Long = 100       ' stop listing after this many problems

' Checks a source file's data BEFORE it touches the workbook.
' Writes each problem to tblExceptions and returns how many were found (0 = file is OK).
Public Function ValidateTransactions(ByVal srcData As Range, ByVal fileName As String) As Long
    Dim required As Variant
    Dim data As Variant
    Dim i As Long
    Dim r As Long
    Dim problems As Long

    ClearExceptions

    ' 1. Required columns, in the expected order
    required = Array("InvoiceNo", "StockCode", "Description", "Quantity", _
                     "InvoiceDate", "UnitPrice", "CustomerID", "Country")
    For i = 0 To 7
        If srcData.Cells(1, i + 1).Value <> required(i) Then
            AddException fileName, 1, CStr(required(i)), CStr(srcData.Cells(1, i + 1).Value), _
                         "Missing or misplaced column"
            problems = problems + 1
        End If
    Next i
    If problems > 0 Then                         ' columns are wrong, so row checks would be meaningless
        ValidateTransactions = problems
        Exit Function
    End If

    ' 2. At least one data row
    If srcData.Rows.Count < 2 Then
        AddException fileName, 1, "", "", "File has no data rows"
        ValidateTransactions = 1
        Exit Function
    End If

    ' 3. Row-by-row checks (on an in-memory copy, for speed)
    data = srcData.Value
    For r = 2 To UBound(data, 1)                 ' UBound(data, 1) = number of rows in the grid
        If IsEmpty(data(r, 4)) Or Not IsNumeric(data(r, 4)) Then
            AddException fileName, r, "Quantity", CStr(data(r, 4)), "Not a number"
            problems = problems + 1
        End If

        If IsEmpty(data(r, 6)) Or Not IsNumeric(data(r, 6)) Then
            AddException fileName, r, "UnitPrice", CStr(data(r, 6)), "Not a number"
            problems = problems + 1
        End If

        If Not IsDate(data(r, 5)) Then
            AddException fileName, r, "InvoiceDate", CStr(data(r, 5)), "Not a valid date"
            problems = problems + 1
        End If

        If problems >= MAX_EXCEPTIONS Then Exit For
    Next r

    ValidateTransactions = problems
End Function

Private Sub ClearExceptions()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Exceptions").ListObjects("tblExceptions")
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
End Sub

Private Sub AddException(ByVal fileName As String, ByVal fileRow As Long, ByVal colName As String, _
                         ByVal badValue As String, ByVal problem As String)
    Dim newRow As ListRow
    Set newRow = ThisWorkbook.Worksheets("Exceptions").ListObjects("tblExceptions").ListRows.Add
    newRow.Range(1, 1).Value = Now
    newRow.Range(1, 2).Value = fileName
    newRow.Range(1, 3).Value = fileRow
    newRow.Range(1, 4).Value = colName
    newRow.Range(1, 5).Value = badValue
    newRow.Range(1, 6).Value = problem
End Sub

