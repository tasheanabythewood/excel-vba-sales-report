Attribute VB_Name = "modMain"
Option Explicit

' One click: import (which validates and refreshes), then export, but only if the import worked
Public Sub RunReport()
    If ImportTransactions() Then ExportReport
End Sub
