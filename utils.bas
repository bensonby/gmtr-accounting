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

Public Function IsFileExist(PathName As String) As Boolean
     'Macro Purpose: Function returns TRUE if the specified file
     '               or folder exists, false if not.
     'PathName     : Supports Windows mapped drives or UNC
     '             : Supports Macintosh paths
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path
     '               Accepts with/without trailing "\" (Windows)
     '               Accepts with/without trailing ":" (Macintosh)
     
    Dim iTemp As Integer
     
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(PathName)
     
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        IsFileExist = True
    Case Else
        IsFileExist = False
    End Select
     
     'Resume error checking
    On Error GoTo 0
End Function
