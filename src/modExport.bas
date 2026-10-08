Attribute VB_Name = "modExport"
Option Explicit

Public Sub ExportReport()
    Dim ws As Worksheet
    Dim outFolder As String
    Dim monthText As String
    Dim pdfPath As String

    On Error GoTo ExportFailed                     ' if anything below fails, jump to ExportFailed

    Set ws = ThisWorkbook.Worksheets("Report")

    ' 1. Build the output folder path and create it if missing
    outFolder = ThisWorkbook.Path & "\output\"
    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder

    ' 2. Turn the selected month into text like "2011-01"
    monthText = Format(ws.Range("SelectedMonth").Value, "yyyy-mm")

    ' 3. Build the full file path
    pdfPath = outFolder & "Sales Report " & monthText & ".pdf"

    ' 4. Export
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, OpenAfterPublish:=True

    MsgBox "Report saved to:" & vbNewLine & pdfPath, vbInformation, "Export complete"
    Exit Sub                                       '  success, so stop here and skip the error section

ExportFailed:                                      ' only reached if something failed
    MsgBox "Couldn't save the PDF. If the file is open in another program, close it and try again" & vbNewLine & vbNewLine & _
           "Details: " & Err.Description, vbCritical, "Export error"
End Sub

