Attribute VB_Name = "ClearData"
Option Explicit

Public Sub ClearWorksheetData()
  Dim dataWorksheetNames() As String
  Dim numberOfWorksheets As Long
  Dim worksheetName As String
  Dim i As Long

  If Config.IsWorksheetNamesValid() = False Then
    Exit Sub
  End If

  dataWorksheetNames = Config.GetDataWorksheetNames()
  numberOfWorksheets = UBound(dataWorksheetNames) + 1 ' zero based
  For i = 0 To numberOfWorksheets - 1
    worksheetName = dataWorksheetNames(i)
    ThisWorkbook.Sheets(worksheetName).UsedRange.ClearContents
  Next i
End Sub
