Attribute VB_Name = "CheckExpiry"

Option Explicit

Private Const EXPIRY_CHECK_URL As String = "https://h3l4cv0j9j.execute-api.us-east-1.amazonaws.com/prod/validate-license?key="
Private Const CODE_INPUT_RANGE_NAME As String = "CODE_INPUT"
Private Const API_KEY As String = "<Insert the API Key here>"
Private Const MAIN_WORKSHEET_NAME As String = "Output"
Private Const PASSWORD As String = "gmtr-Oursky"

Private Function getFirstDataWorksheet() As String
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> MAIN_WORKSHEET_NAME Then
      getFirstDataWorksheet = ws.Name
      Exit For
    End If
  Next ws
  getFirstDataWorksheet = ""
End Function

Public Sub OnWorkbookOpen()
  Dim firstDataWorksheet As String
  Dim code As String
  Dim isSpreadsheetValid As Boolean

  'Hide the main worksheet. Only show after the expiry date check is passed
  'Since we cannot hide the last worksheet, we will unhide the first data worksheet
  firstDataWorksheet = getFirstDataWorksheet()
  ThisWorkbook.Unprotect PASSWORD
  ThisWorkbook.Sheets(firstDataWorksheet).Visible = True
  ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Visible = xlVeryHidden
  ThisWorkbook.Protect PASSWORD, True, False

  'Initial Check
  code = GetCode()
  isSpreadsheetValid = CheckExpiryDate(code)

  If code = vbNullString Then
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub
  End If

  'Subsequent prompts if first check fails
  Do While isSpreadsheetValid = False
    code = PromptCode()
    If code = vbNullString Then
      ThisWorkbook.Close SaveChanges:=False
      Exit Sub
    End If
    isSpreadsheetValid = CheckExpiryDate(code)
  Loop

  'Show main worksheet
  ThisWorkbook.Unprotect PASSWORD
  ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Visible = True
  ThisWorkbook.Sheets(firstDataWorksheet).Visible = False
  ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Activate
  ThisWorkbook.Protect PASSWORD, True, False
End Sub

Private Function GetCode() As String
  Dim currentValue As String
  currentValue = ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Range(CODE_INPUT_RANGE_NAME).Value
  If currentValue <> "" Then
    GetCode = currentValue
  Else
    GetCode = PromptCode()
  End If
End Function

Private Function PromptCode() As String
  Dim code As String
  code = InputBox("Enter code:")
  ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Range(CODE_INPUT_RANGE_NAME).Value = code
  PromptCode = code
End Function

Private Function CheckExpiryDate(code As String) As Boolean
  Dim result As String
  Dim WinHttpReq As Object

  If code = vbNullString Then
    CheckExpiryDate = False
    Exit Function
  End If

  Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  WinHttpReq.Open "GET", EXPIRY_CHECK_URL & code, False
  WinHttpReq.setRequestHeader "x-api-key", API_KEY
  WinHttpReq.send

  If WinHttpReq.Status = 200 Then
    result = WinHttpReq.responseText
    If result = """true""" Then
      CheckExpiryDate = True
    Else
      If result = """false""" Then
        MsgBox "Code Expired."
      ElseIf result = """notfound""" Then
        MsgBox "Code Invalid."
      Else
        MsgBox "Unknown Error: " & result
      End If
      CheckExpiryDate = False
    End If
  Else
    MsgBox "An error has occurred. Please check your internet connection. Status: " & CStr(WinHttpReq.Status)
    CheckExpiryDate = False
  End If
End Function
