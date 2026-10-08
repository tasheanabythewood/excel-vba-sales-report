Attribute VB_Name = "modReport"
Option Explicit

' Returns the month (as the 1st of the month) found in tblTransactions
Public Function GetDataMonth() As Date
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Transactions").ListObjects("tblTransactions")
    GetDataMonth = Application.WorksheetFunction.Min(tbl.ListColumns("SalesMonth").DataBodyRange)
End Function

' Points the whole report at one month
Public Sub RefreshReport(ByVal reportMonth As Date)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim monthKey As String

    Set ws = ThisWorkbook.Worksheets("Report")
    Set pt = ws.PivotTables("ptCountry")

    ' 0. Stop early if the month isn't in the data, before changing anything
    monthKey = Format(reportMonth, "yyyy-mm")

    If Application.WorksheetFunction.CountIfs( _
        ThisWorkbook.Worksheets("Transactions").ListObjects("tblTransactions") _
            .ListColumns("MonthKey").DataBodyRange, monthKey) = 0 Then
        MsgBox "No data for " & Format(reportMonth, "mmm yyyy") & "." & vbNewLine & _
               "Import that month's file first.", vbExclamation, "Month not loaded"
        Exit Sub
    End If

    ' 1. Point the KPI formulas at the month
    ws.Range("SelectedMonth").Value = reportMonth

    ' 2. Recalculate so the formulas pick up the new month
    Application.Calculate

    ' 3. Re-read the table into the PivotTable, dropping months that no longer exist
    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
    pt.PivotCache.Refresh

    ' 4. Filter the PivotTable to the month
    pt.PivotFields("MonthKey").ClearAllFilters
    pt.PivotFields("MonthKey").CurrentPage = monthKey

    FormatReport                                  ' re-apply formatting and column widths
End Sub

' Button-friendly version: refreshes whatever month the user typed into SelectedMonth
Public Sub RefreshSelectedMonth()
    Dim wanted As Date
    wanted = ThisWorkbook.Worksheets("Report").Range("SelectedMonth").Value
    ThisWorkbook.Worksheets("Report").Range("SelectedMonth").Value = GetDataMonth()   ' reset to what's loaded
    RefreshReport wanted                                                              ' only changes it if valid
End Sub

