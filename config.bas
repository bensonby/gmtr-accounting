Attribute VB_Name = "Config"
Option Explicit

Public Const LOCAL_DATA_FILENAME As String = "data.json"
Public Const WORKSHEET_NAMES = "MLR,1,1a,1b,2,2a,2b,3,3a,3b,4,4a,4b,5,5a,5b,6,6a,6b,7,7a,7b,8,8a,8b,9,9a,9b,10,10a,10b,11,11a,11b,12,12a,12b,13,13a,13b,14,14a,14b,15,15a,15b,16,16a,16b,17,17a,17b,18,18a,18b,19,19a,19b,20,20a,20b,21,21a,21b,22,22a,22b,23,23a,23b,24,24a,24b,25,25a,25b,26,26a,26b,27,27a,27b,28,28a,28b,29,29a,29b,30,30a,30b,31,31a,31b,32,32a,32b,33,33a,33b,34,34a,34b,35,35a,35b,36,36a,36b,37,38,39,40,41,41a,41b,42,43,43a,43b,44,44a,44b,45,45a,45b,46,46a,46b,47,47a,47b,48,49,49a,49b,50,51,52,53,54,55,56,57,58,58a,58b,59,59a,59b,60,60a,60b,61,61a,61b,62,62a,62b,63,63a,63b,64,64a,64b,65,65a,65b,66,67,67a,67b,68,68a,68b,69,69a,69b,70,70a,70b,71,71a,71b,72,72a,72b"
Public Const EXPIRY_DATE = #12/21/2017#

Public Function IsExpired() As Boolean
  If Date() >= EXPIRY_DATE Then
    IsExpired = True
  Else
    IsExpired = False
  End If
End Function

Public Function GetDataWorksheetNames() As String()
  GetDataWorksheetNames = Split(WORKSHEET_NAMES, ",")
End Function

Public Function IsWorksheetNamesValid() As Boolean
  Dim worksheets() As String
  Dim i As Long

  worksheets = GetDataWorksheetNames()
  For i = 0 To UBound(worksheets)
    If Not Utils.IsWorksheetExist("", worksheets(i)) Then
      MsgBox "Worksheet not found: " & worksheets(i)
      IsWorksheetNamesValid = False
      Exit Function
    End If
  Next i
  IsWorksheetNamesValid = True
End Function
