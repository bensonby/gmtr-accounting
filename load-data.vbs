' Dependency: JsonConverter.bas, from https://github.com/VBA-tools/VBA-JSON
' Also enable Microsoft Scripting Runtime from Tools -> Refernece in VBA window

' Note: Array in the JSON Objects or Collections starts at 1 not 0

Option Explicit

Public Const FOR_READING As Integer = 1

Public Sub LoadJsonFromFile()
  Dim fs As Object
  Dim fsTextStream As Object
  Dim jsonData As Dictionary
  Dim worksheetName As Variant ' need to be Variant not String, for parsing from JSON
  Dim worksheetData As Variant
  Dim startCell As Range
  Dim endCell As Range

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set fsTextStream = fs.OpenTextFile("data.json", FOR_READING)
  Set jsonData = JsonConverter.ParseJson(fsTextStream.ReadAll())

  For Each worksheetName in jsonData.Keys()
    Application.StatusBar = "Loading data for " & CStr(worksheetName)

    worksheetData = ConvertCollectionToArray(jsonData(worksheetName))

    Set startCell = ThisWorkbook.Sheets(worksheetName).Cells(1, 1)
    Set endCell = ThisWorkbook.Sheets(worksheetName).Cells(UBound(worksheetData) + 1, UBound(worksheetData, 2) + 1)
    Range(startCell, endCell).Value = worksheetData

    ' To interpret all numbers correctly
    Range(startCell, endCell).Value = Range(startCell, endCell).Value
  Next

  MsgBox "Completed"

  ' Restore environment
  Application.StatusBar = False
End Sub

Private Function ConvertCollectionToArray(data As Collection) As String()
  Dim result() As String
  Dim rowCount As Long
  Dim columnCount As Long
  Dim i As Long
  Dim j As Long

  rowCount = data.Count
  columnCount = data(1).Count

  ReDim result(rowCount - 1, columnCount - 1) ' zero based

  For i = 1 To rowCount
    For j = 1 To columnCount
      result(i - 1, j - 1) = data(i)(j)
    Next j
  Next i

  ConvertCollectionToArray = result
End Function
