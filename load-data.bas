Attribute VB_Name = "LoadData"
' Dependency: JsonConverter.bas, from https://github.com/VBA-tools/VBA-JSON
' Also enable Microsoft Scripting Runtime from Tools -> Refernece in VBA window

' Note: Array in the JSON Objects or Collections starts at 1 not 0

Option Explicit

Public Const FOR_READING As Integer = 1
Public Const JSON_URL As String = "https://s3-ap-southeast-1.amazonaws.com/gmtresearch-accounting-screen/data.json"

Public Sub OnWorkbookOpen()
  If Utils.IsFileExist(ThisWorkbook.Path & "\" & Config.LOCAL_DATA_FILENAME) = False Then
    Call DownloadAndLoadJsonFile
  Else
    Call LoadJsonFromFile
  End If
End Sub

Public Sub DownloadAndLoadJsonFile()
  If DownloadJsonFile() = True Then
    Call LoadJsonFromFile
  End If
End Sub

Public Function DownloadJsonFile() As Boolean
  Dim WinHttpReq As Object
  Dim oStream As Object

  If Config.IsExpired() = True Then
    MsgBox "Worksheet expired!"
    DownloadJsonFile = False
    Exit Function
  End If

  Application.StatusBar = "Downloading data file..."

  Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  WinHttpReq.Open "GET", JSON_URL, False
  WinHttpReq.send

  If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile ThisWorkbook.Path & "\" & Config.LOCAL_DATA_FILENAME, 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
    Application.StatusBar = "Download complete! Loading data..."
    DownloadJsonFile = True
  Else
    MsgBox "Download Error! HTTP Status: " & CStr(WinHttpReq.Status)
    Application.StatusBar = False
    DownloadJsonFile = False
  End If
End Function

Private Sub LoadJsonFromFile()
  Dim fs As Object
  Dim fsTextStream As Object
  Dim jsonData As Dictionary
  Dim worksheetName As Variant ' need to be Variant not String, for parsing from JSON
  Dim worksheetData As Variant
  Dim startCell As Range
  Dim endCell As Range

  If Config.IsExpired() = True Then
    MsgBox "Worksheet expired!"
    Exit Sub
  End If

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set fsTextStream = fs.OpenTextFile(ThisWorkbook.Path & "\" & Config.LOCAL_DATA_FILENAME, FOR_READING)
  Set jsonData = JsonConverter.ParseJson(fsTextStream.ReadAll())

  For Each worksheetName in jsonData.Keys()
    If Not Utils.IsWorksheetExist("", CStr(worksheetName)) Then
      MsgBox "Worksheet not found: " & worksheetName
      Exit Sub
    End If
  Next

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
