Attribute VB_Name = "ExportData"
Option Explicit

Public Const MAX_NUMBER_OF_ROWS As Long = 200
Public Const MAX_NUMBER_OF_COLUMNS As Long = 300

Public Sub SaveJsonToFile()
  Dim fs As Object
  Dim jsonFile As Object
  Dim data As String

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set jsonFile = fs.CreateTextFile(ThisWorkbook.Path & "\" & "data.json")

  If Config.IsWorksheetNamesValid() = False Then
    Exit Sub
  End If

  data = ExportDataAsJson()

  Application.StatusBar = "Writing data to file..."
  jsonFile.writeLine(data)

  MsgBox "Completed"
  Application.StatusBar = False
End Sub

Private Function ExportDataAsJson()
  Dim result() As String
  Dim dataWorksheetNames() As String
  Dim numberOfWorksheets As Long
  Dim worksheetName As String
  Dim worksheetJson As String
  Dim i As Long

  dataWorksheetNames = Config.GetDataWorksheetNames()
  numberOfWorksheets = UBound(dataWorksheetNames) + 1 ' zero based
  ReDim result(numberOfWorksheets - 1) ' zero based

  For i = 0 To numberOfWorksheets - 1
    worksheetName = dataWorksheetNames(i)
    Application.StatusBar = "Parsing data for worksheet " & worksheetName
    worksheetJson = ExportWorksheetAsJson(worksheetName)
    result(i) = """" & worksheetName & """: " & worksheetJson
  Next i

  ExportDataAsJson = "{" & Join(result, ",") & "}"
End Function

Private Function ExportWorksheetAsJson(worksheetName As String)
  Dim result() As String
  Dim numberOfRows As Long
  Dim usedRange As Range
  Dim rowCount As Long
  Dim columnCount As Long
  Dim i As Long

  Set usedRange = ThisWorkbook.Sheets(worksheetName).UsedRange
  rowCount = WorksheetFunction.Min(usedRange.Rows.Count, MAX_NUMBER_OF_ROWS)
  columnCount = WorksheetFunction.Min(usedRange.Columns.Count, MAX_NUMBER_OF_COLUMNS)

  ReDim result(rowCount - 1) ' zero based

  For i = 1 To rowCount
    result(i - 1) = ExportRowAsJson(usedRange.Rows(i), columnCount)
  Next i

  ExportWorksheetAsJson = "[" & Join(result, ",") & "]"
End Function

Private Function ExportRowAsJson(row As Range, columnCount As Long) As String
  Dim result() As String
  Dim i As Long

  ReDim result(columnCount - 1) ' zero based

  For i = 1 To columnCount
    result(i - 1) = ExportCellAsJson(row.Cells(i))
  Next i
  ExportRowAsJson = "[" & Join(result, ",") & "]"
End Function

Private Function ExportCellAsJson(cell As Range) As String
  ExportCellAsJson = """" & Replace(CStr(cell.Value2), """", "\""") & """"
End Function
