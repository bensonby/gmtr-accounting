Attribute VB_Name = "Utils"
Option Explicit

Public Function IsWorksheetExist(wb As String, ws As String) As Boolean
  Dim wSheet As Worksheet
  If wb = "" Then
    wb = ThisWorkbook.name
  End If
  On Error Resume Next
  Set wSheet = Workbooks(wb).Worksheets(ws)
  If wSheet Is Nothing Then
    On Error GoTo 0
    IsWorksheetExist = False
  Else
    Set wSheet = Nothing
    On Error GoTo 0
    IsWorksheetExist = True
  End If
End Function
