Attribute VB_Name = "CheckExpiry"

Option Explicit

Private Const EXPIRY_CHECK_URL As String = "https://h3l4cv0j9j.execute-api.us-east-1.amazonaws.com/prod/validate-license?key="
Private Const CODE_INPUT_RANGE_NAME As String = "CODE_INPUT"
Private Const API_KEY As String = "<Insert the API Key here>"
Private Const MAIN_WORKSHEET_NAME As String = "Output"
Private Const LOADING_WORKSHEET_NAME As String = "Loading"
Private Const PASSWORD As String = "gmtr-Oursky"

Public Sub OnWorkbookOpen()
  Dim ws As Worksheet
  Dim code As String
  Dim isSpreadsheetValid As Boolean

  'Change macro security settings to lowest for worksheet manipulation
  'Will restore at end of script

  Dim originalSecurity As Long
  originalSecurity = Application.AutomationSecurity
  Application.AutomationSecurity = msoAutomationSecurityLow

  'Show only the Loading Worksheet, hide all others
  'Will hide the loading worksheet and show the main worksheet after expiry date check is passed
  ThisWorkbook.Unprotect PASSWORD
  ThisWorkbook.Sheets(LOADING_WORKSHEET_NAME).Visible = True
  For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> LOADING_WORKSHEET_NAME And ws.Visible <> False Then
      ' User cannot see these worksheets even in Unhide menu
      ws.Visible = False
    End If
  Next ws
  ThisWorkbook.Protect PASSWORD, True, False

  'Initial Check
  code = GetCode()
  isSpreadsheetValid = CheckExpiryDate(code)

  If code = vbNullString Then
    Application.AutomationSecurity = originalSecurity
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub
  End If

  'Subsequent prompts if first check fails
  Do While isSpreadsheetValid = False
    code = PromptCode()
    If code = vbNullString Then
      Application.AutomationSecurity = originalSecurity
      ThisWorkbook.Close SaveChanges:=False
      Exit Sub
    End If
    isSpreadsheetValid = CheckExpiryDate(code)
  Loop

  'Show main worksheet
  ThisWorkbook.Unprotect PASSWORD
  ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Visible = True
  ThisWorkbook.Sheets(LOADING_WORKSHEET_NAME).Visible = xlVeryHidden
  ThisWorkbook.Sheets(MAIN_WORKSHEET_NAME).Activate
  ThisWorkbook.Protect PASSWORD, True, False

  Application.AutomationSecurity = originalSecurity
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
