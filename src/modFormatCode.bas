Attribute VB_Name = "modFormatCode"
Option Explicit

Public Sub FormatReport()
    Dim ws As Worksheet
    Dim gbpFormat As String

    Set ws = ThisWorkbook.Worksheets("Report")

    ' ChrW(163) is the pound sign. Building it in code keeps the exported .bas file
    ' plain text, so it displays correctly on GitHub.
    gbpFormat = ChrW(163) & "#,##0.00;[Red]-" & ChrW(163) & "#,##0.00"

    With ws.Range("ReportTitle")
        .Value = "Monthly Sales Report"
        .Font.Size = 20
        .Font.Bold = True
    End With

    ws.Range("KPIMoney").NumberFormat = gbpFormat
    ws.Range("KPILabels").Font.Bold = True
    ws.Range("KPIBlock").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    ws.PivotTables("ptCountry").DataFields(1).NumberFormat = gbpFormat

        ' Stop the PivotTable resizing columns on every refresh; we size them ourselves
    ws.PivotTables("ptCountry").HasAutoFormat = False

    ' Fit columns to the KPI block AND the PivotTable together (not the title in A1)
    ws.Range(ws.Range("KPIBlock"), ws.PivotTables("ptCountry").TableRange2).Columns.AutoFit
End Sub
