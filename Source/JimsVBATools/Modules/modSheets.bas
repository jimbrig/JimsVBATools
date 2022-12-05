Attribute VB_Name = "modSheets"
' @Folder("Worksheets")
Option Explicit

Private Const MaxLengthSheetName As Integer = 31

Public Sub AddNewWorksheet()
' Attribute AddNewWorksheet.VB_ProcData.VB_Invoke_Func = "w\n14"

  Const cstrTitle As String = "Add new worksheet"
  Const cstrPrompt As String = "Give the name for the new worksheet." & vbCrLf & "Not allowed are the characters: : \ / ? * [ and ]"

  Dim strInput As String
  Dim strDefault As String: strDefault = "Sheet" 'setting initial value for inputbox can be useful'
  Dim strInputErrorMessage As String
  Dim defaultColor As String: defaultColor = ""
  Dim booValidatedOk As Boolean: booValidatedOk = False
  
  On Error GoTo HandleError
    
    Do
        strInput = InputBox(Prompt:=cstrPrompt, Title:=cstrTitle, Default:=strDefault)
        If Len(strInput) = 0 Then GoTo HandleExit
        GoSub ValidateInput
        If Not booValidatedOk Then
            If vbCancel = MsgBox(strInputErrorMessage & "Retry?", vbExclamation + vbOKCancel) Then GoTo HandleExit
        End If
    Loop While Not booValidatedOk
        
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim shts As Sheets: Set shts = wb.Sheets
    Dim obj As Object
    Set obj = shts.Add(Before:=ActiveSheet, Count:=1, Type:=XlSheetType.xlWorksheet)
    obj.Name = strInput
    
HandleExit:
    Exit Sub
HandleError:
    MsgBox Err.Description
    Resume HandleExit
    
ValidateInput:
    If SheetExists(strSheetName:=strInput) Then
        strInputErrorMessage = "Sheet already exists. "
    ElseIf Not IsValidSheetName(strSheetName:=strInput) Then
        strInputErrorMessage = "Sheetname not allowed. "
    Else
        booValidatedOk = True
    End If
    Return
End Sub

