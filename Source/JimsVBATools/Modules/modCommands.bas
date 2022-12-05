Attribute VB_Name = "modCommands"
' @Folder("Commands")
Option Explicit

' ----------------------------
' Declarations
' ----------------------------

' ----------------------------
' ExePath Variable Declaration
' ----------------------------
Dim ExePath As String

' ----------------------------
' Custom "Program" Type
' ----------------------------
Type Program
  Name As String
  Path As String
  Args As String
  Message As String
  WindowStyle As VbAppWinStyle
  ProcessID As Double
End Type

' ----------------------------
' TEMPLATE
' ----------------------------
' Public Sub cmdOpenProgram()
'
'   Dim Prog As Program
'
'   ExePath = "<Path to the Executable>"
'
'   With Prog
'     .Name = "OfficeRibbonXEditor"
'     .Path = ExePath
'     .Args = ""
'     .WindowStyle = vbNormalFocus
'     .ProcessID = Shell(.Path, .WindowStyle)
'     .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
'   End With
'
'   Debug.Print Prog.Message
'
' End Sub
' -----------------------------

' -----------------------------
' OfficeRibbonXEditor
' -----------------------------
Public Sub CmdOpenRibbonXEditor()
     
    Dim Prog As Program
    
    ExePath = "C:\Users\jbriggs010\AppData\Local\OfficeRibbonEditor\OfficeRibbonXEditor.exe"
    
    With Prog
      .Name = "OfficeRibbonXEditor"
      .Path = ExePath
      .Args = ""
      .WindowStyle = vbNormalFocus
      .ProcessID = Shell(.Path, .WindowStyle)
      .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
    End With
        
    Debug.Print Prog.Message

End Sub

' -----------------------------
' Windows Terminal
' -----------------------------
Public Sub CmdOpenWindowsTerminal()

  Dim Prog As Program
    
  ExePath = "C:\Users\jbriggs010\AppData\Local\Microsoft\WindowsApps\wt.exe"
    
  With Prog
    .Name = "Windows Terminal (Preview)"
    .Path = ExePath
    .Args = ""
    .WindowStyle = vbNormalFocus
    .ProcessID = Shell(.Path, .WindowStyle)
    .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
  End With
        
  Debug.Print Prog.Message
  
End Sub

' -----------------------------
' VSCode (Insiders)
' -----------------------------
Public Sub CmdOpenVSCode()

  Dim Prog As Program
    
  ExePath = "C:\Users\jbriggs010\AppData\Local\Microsoft\WindowsApps\wt.exe"
    
  With Prog
    .Name = "Windows Terminal (Preview)"
    .Path = ExePath
    .Args = ""
    .WindowStyle = vbNormalFocus
    .ProcessID = Shell(.Path, .WindowStyle)
    .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
  End With
        
  Debug.Print Prog.Message
  
End Sub

' -----------------------------
' Notepad
' -----------------------------
Public Sub CmdOpenNotepad()

  Dim Prog As Program
    
  ExePath = "C:\Users\jbriggs010\AppData\Local\Microsoft\WindowsApps\wt.exe"
    
  With Prog
    .Name = "Windows Terminal (Preview)"
    .Path = ExePath
    .Args = ""
    .WindowStyle = vbNormalFocus
    .ProcessID = Shell(.Path, .WindowStyle)
    .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
  End With
        
  Debug.Print Prog.Message
  
End Sub

' -----------------------------
' DAXStudio
' -----------------------------
Public Sub CmdOpenDAXStudio()

  Dim Prog As Program
    
  ExePath = "C:\Users\jbriggs010\AppData\Local\Microsoft\WindowsApps\wt.exe"
    
  With Prog
    .Name = "Windows Terminal (Preview)"
    .Path = ExePath
    .Args = ""
    .WindowStyle = vbNormalFocus
    .ProcessID = Shell(.Path, .WindowStyle)
    .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
  End With
        
  Debug.Print Prog.Message
  
End Sub

' -----------------------------
' Explorer
' -----------------------------
Public Sub CmdOpenExplorer()

  Dim Prog As Program
    
  ExePath = "C:\Users\jbriggs010\AppData\Local\Microsoft\WindowsApps\wt.exe"
    
  With Prog
    .Name = "Windows Terminal (Preview)"
    .Path = ExePath
    .Args = ""
    .WindowStyle = vbNormalFocus
    .ProcessID = Shell(.Path, .WindowStyle)
    .Message = "Launched  " & .Name & " Program with the PID: " & .ProcessID
  End With
        
  Debug.Print Prog.Message
  
End Sub


