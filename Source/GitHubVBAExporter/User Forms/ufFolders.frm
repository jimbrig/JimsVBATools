VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFolders 
   Caption         =   "JKP-ADS VBA Project Exporter settings"
   ClientHeight    =   11070
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8445.001
   OleObjectBlob   =   "ufFolders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NoEvents As Boolean

Public OK As Boolean
Public WorkbookName As String
Public MainPath As String

Public UserformsFolder As String
Public DeleteExisting As Boolean
Public ExcelObjectsFolder As String
Public ModulesFolder As String
Public ClassModulesFolder As String
Public RibbonXFolder As String
Public AddInFolder As String
Public ExcelFolder As String
Public CopyExcelFile As Boolean
Public CreateAddIn As Boolean
Public ExportRibbonX As Boolean

Private Const MCMINWIDTH As Double = 360

Private mdStartX As Double
Private mbSizing As Boolean

Public Sub Initialize()
    Dim vbp As VBProject
    Dim oCtl As MSForms.control
    Dim oWb As Workbook
    Dim oAddin As AddIn
    For Each oCtl In Me.Controls
        If TypeName(oCtl) = "TextBox" Then
            oCtl.Value = GetSetting(GCSAPPREGKEY, "Settings", oCtl.Tag, oCtl.Tag)
        ElseIf TypeName(oCtl) = "CheckBox" Then
            oCtl.Value = CBool(GetSetting(GCSAPPREGKEY, "Settings", oCtl.Tag, 0))
        End If
    Next
    tbxExportLocation.Value = GetSetting(GCSAPPREGKEY, "Settings", "MainPath", "")
    NoEvents = True
    With cbbFile2Export
        On Error Resume Next
        For Each vbp In Application.VBE.VBProjects
            If vbp.Protection = vbext_pp_none Then
                .AddItem Right(vbp.FileName, InStrRev(vbp.FileName, Application.PathSeparator) + 1)
                .List(.ListCount - 1, 1) = vbp.FileName
            Else
                lblProtected.Caption = "One or more protected VBprojects detected, these are not included in the drop-down"
            End If
        Next
    End With
    NoEvents = False
    If cbbFile2Export.ListCount = 1 Then cbbFile2Export.ListIndex = 0
    HandleControls
End Sub

Private Sub HandleControls()
    Dim bOK As Boolean
    bOK = True
    If cbbFile2Export.ListIndex = -1 Then
        bOK = False
    ElseIf Len(tbxExportLocation.Value) = 0 Then
        bOK = False
    End If
    cmbOK.Enabled = bOK
    tbxExcelFile.Enabled = cbxExcelFile.Value
    lblExcelFile.Enabled = cbxExcelFile.Value
    tbxAddinFile.Enabled = cbxCreateAddIn.Value
    lblAddinFile.Enabled = cbxCreateAddIn.Value
    tbxRibbonX.Enabled = cbxRibbonX.Value
    lblRibbonX.Enabled = cbxRibbonX.Value
End Sub

Private Sub cbxCreateAddIn_Click()
    CreateAddIn = cbxCreateAddIn.Value
    HandleControls
End Sub

Private Sub cbxDeleteExisting_Click()
    DeleteExisting = cbxDeleteExisting.Value
    If DeleteExisting Then
        cbxDeleteExisting.ForeColor = vbRed
    Else
        cbxDeleteExisting.ForeColor = &H80000015
    End If
    HandleControls
End Sub

Private Sub cbxExcelFile_Click()
    CopyExcelFile = cbxExcelFile.Value
    HandleControls
End Sub

Private Sub cbxRibbonX_Click()
    ExportRibbonX = cbxRibbonX.Value
    HandleControls
End Sub

Private Sub cmbCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub cmbMainFolder_Click()
    Dim oFD As FileDialog
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    With oFD
        .Title = "Select destination folder"
        If .Show Then
            tbxExportLocation.Value = .SelectedItems(1)
        Else
            tbxExportLocation.Value = vbNullString
        End If
    End With
End Sub

Private Sub cmbOK_Click()
    Dim oCtl As MSForms.control
    OK = True
    For Each oCtl In Me.Controls
        If TypeName(oCtl) = "TextBox" Then
            SaveSetting GCSAPPREGKEY, "Settings", oCtl.Tag, oCtl.Value
        ElseIf TypeName(oCtl) = "CheckBox" Then
            If oCtl.Value Then
                SaveSetting GCSAPPREGKEY, "Settings", oCtl.Tag, "1"
            Else
                SaveSetting GCSAPPREGKEY, "Settings", oCtl.Tag, "0"
            End If
        End If
    Next
    SaveSetting GCSAPPREGKEY, "Settings", "MainPath", MainPath
    Me.Hide
End Sub

Private Sub cbbFile2Export_Change()
    If NoEvents Then Exit Sub
    If cbbFile2Export.ListIndex <> -1 Then
        WorkbookName = cbbFile2Export.Value
    End If
    HandleControls
End Sub

Private Sub lblSizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mbSizing = True
    mdStartX = X
End Sub

Private Sub lblSizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim oCtl As control
    If mbSizing Then
        If Me.Width + (X - mdStartX) > MCMINWIDTH Then
            Me.Width = Me.Width + (X - mdStartX)
            lblSizer.Left = lblSizer.Left + (X - mdStartX)
            For Each oCtl In Me.Controls
                Select Case LCase(TypeName(oCtl))
                Case "textbox", "frame", "combobox"
                    oCtl.Width = oCtl.Width + (X - mdStartX)
                Case "commandbutton"
                    oCtl.Left = oCtl.Left + (X - mdStartX)
                End Select
            Next
        End If
    End If
End Sub

Private Sub lblSizer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mbSizing = False
End Sub

Private Sub tbxAddinFile_Change()
    HandleControls
    AddInFolder = tbxAddinFile.Value
End Sub

Private Sub tbxClassModules_Change()
    HandleControls
    ClassModulesFolder = tbxClassModules.Value
End Sub

Private Sub tbxExcelFile_Change()
    HandleControls
    ExcelFolder = tbxExcelFile.Value
End Sub

Private Sub tbxExcelObjects_Change()
    HandleControls
    ExcelObjectsFolder = tbxExcelObjects.Value
End Sub

Private Sub tbxExportLocation_Change()
    HandleControls
    If Right(tbxExportLocation.Value, 1) <> Application.PathSeparator Then
        tbxExportLocation.Value = tbxExportLocation.Value & Application.PathSeparator
    End If
    MainPath = tbxExportLocation.Value
End Sub

Private Sub tbxModules_Change()
    HandleControls
    ModulesFolder = tbxModules.Value
End Sub

Private Sub tbxRibbonX_Change()
    HandleControls
    RibbonXFolder = tbxRibbonX.Value
End Sub

Private Sub tbxUserforms_Change()
    HandleControls
    UserformsFolder = tbxUserforms.Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    cmbCancel_Click
End Sub
