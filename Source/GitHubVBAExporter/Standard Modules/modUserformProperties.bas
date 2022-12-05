Attribute VB_Name = "modUserformProperties"
Option Explicit

Sub ListForms(oWb As Workbook, sExportPath As String)
    Dim oVBProj As VBProject
    Dim cControl As control
    Dim oComp As VBComponent
    Dim oSh As Worksheet
    Dim lCount As Long
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    On Error GoTo LocErr
    Set oVBProj = oWb.VBProject
    ThisWorkbook.Worksheets("Userforms").Copy
    Set oSh = ActiveSheet
    For Each oComp In oVBProj.VBComponents
        If oComp.Type = vbext_ct_MSForm Then
            Application.StatusBar = oComp.Name
            ListProperties oComp.Name, oComp.Designer, oSh, lCount
            lCount = lCount + 1
            For Each cControl In oComp.Designer.Controls
                ListProperties oComp.Name, cControl, oSh, lCount
                lCount = lCount + 1
            Next
        End If
    Next
    If Right(sExportPath, 1) <> Application.PathSeparator Then sExportPath = sExportPath & Application.PathSeparator
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs sExportPath & "FormsAndControlsProperties.txt", xlCSV
    ActiveWorkbook.Close False
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
LocErr:
    MsgBox Err.Description
    Stop
    Resume 'Next
End Sub

Sub ListProperties(ByVal sFormName As String, ByVal cControl As Object, oSh As Worksheet, lCount As Long)
    On Error Resume Next
    With oSh
        .Names("Userforms.Form").RefersToRange.Offset(lCount, 0) = sFormName
        .Names("Userforms.Controlname").RefersToRange.Offset(lCount, 0) = cControl.Name
        .Names("Userforms.NewName").RefersToRange.Offset(lCount, 0) = cControl.Name    'Remember name
        .Names("Userforms.Type").RefersToRange.Offset(lCount, 0) = TypeName(cControl)
        .Names("Userforms.Accelerator").RefersToRange.Offset(lCount, 0) = cControl.Accelerator
        .Names("Userforms.ActiveControl").RefersToRange.Offset(lCount, 0) = cControl.ActiveControl
        .Names("Userforms.Alignment").RefersToRange.Offset(lCount, 0) = cControl.Alignment
        .Names("Userforms.AutoSize").RefersToRange.Offset(lCount, 0) = cControl.AutoSize
        .Names("Userforms.BackColor").RefersToRange.Offset(lCount, 0) = cControl.BackColor
        .Names("Userforms.BackStyle").RefersToRange.Offset(lCount, 0) = cControl.BackStyle
        .Names("Userforms.BorderColor").RefersToRange.Offset(lCount, 0) = cControl.BorderColor
        .Names("Userforms.BorderStyle").RefersToRange.Offset(lCount, 0) = cControl.BorderStyle
        .Names("Userforms.BoundColumn").RefersToRange.Offset(lCount, 0) = cControl.BoundColumn
        .Names("Userforms.BoundValue").RefersToRange.Offset(lCount, 0) = cControl.BoundValue
        .Names("Userforms.Cancel").RefersToRange.Offset(lCount, 0) = cControl.Cancel
        .Names("Userforms.CanPaste").RefersToRange.Offset(lCount, 0) = cControl.CanPaste
        .Names("Userforms.CanRedo").RefersToRange.Offset(lCount, 0) = cControl.CanRedo
        .Names("Userforms.CanUndo").RefersToRange.Offset(lCount, 0) = cControl.CanUndo
        .Names("Userforms.Caption").RefersToRange.Offset(lCount, 0) = cControl.Caption
        .Names("Userforms.Column").RefersToRange.Offset(lCount, 0) = cControl.Column
        .Names("Userforms.ColumnCount").RefersToRange.Offset(lCount, 0) = cControl.ColumnCount
        .Names("Userforms.ColumnHeads").RefersToRange.Offset(lCount, 0) = cControl.ColumnHeads
        .Names("Userforms.ColumnWidths").RefersToRange.Offset(lCount, 0) = cControl.ColumnWidths
        .Names("Userforms.ControlSource").RefersToRange.Offset(lCount, 0) = cControl.ControlSource
        .Names("Userforms.ControlTipText").RefersToRange.Offset(lCount, 0) = cControl.ControlTipText
        .Names("Userforms.Cycle").RefersToRange.Offset(lCount, 0) = cControl.Cycle
        .Names("Userforms.Default").RefersToRange.Offset(lCount, 0) = cControl.Default
        .Names("Userforms.DrawBuffer").RefersToRange.Offset(lCount, 0) = cControl.DrawBuffer
        .Names("Userforms.Enabled").RefersToRange.Offset(lCount, 0) = cControl.Enabled
        .Names("Userforms.FontName").RefersToRange.Offset(lCount, 0) = cControl.Font.Name
        .Names("Userforms.FontSize").RefersToRange.Offset(lCount, 0) = cControl.Font.Size
        .Names("Userforms.ForeColor").RefersToRange.Offset(lCount, 0) = cControl.ForeColor
        .Names("Userforms.GroupName").RefersToRange.Offset(lCount, 0) = cControl.GroupName
        .Names("Userforms.Height").RefersToRange.Offset(lCount, 0) = cControl.Height
        .Names("Userforms.HelpContextID").RefersToRange.Offset(lCount, 0) = cControl.HelpContextID
        .Names("Userforms.IMEMode").RefersToRange.Offset(lCount, 0) = cControl.IMEMode
        .Names("Userforms.InsideHeight").RefersToRange.Offset(lCount, 0) = cControl.InsideHeight
        .Names("Userforms.InsideWidth").RefersToRange.Offset(lCount, 0) = cControl.InsideWidth
        .Names("Userforms.IntegralHeight").RefersToRange.Offset(lCount, 0) = cControl.IntegralHeight
        .Names("Userforms.KeepScrollBarsVisible").RefersToRange.Offset(lCount, 0) = cControl.KeepScrollBarsVisible
        .Names("Userforms.LayoutEffect").RefersToRange.Offset(lCount, 0) = cControl.LayoutEffect
        .Names("Userforms.Left").RefersToRange.Offset(lCount, 0) = cControl.Left
        .Names("Userforms.List").RefersToRange.Offset(lCount, 0) = cControl.List
        .Names("Userforms.ListCount").RefersToRange.Offset(lCount, 0) = cControl.ListCount
        .Names("Userforms.ListIndex").RefersToRange.Offset(lCount, 0) = cControl.ListIndex
        .Names("Userforms.ListStyle").RefersToRange.Offset(lCount, 0) = cControl.ListStyle
        .Names("Userforms.Locked").RefersToRange.Offset(lCount, 0) = cControl.Locked
        .Names("Userforms.MatchEntry").RefersToRange.Offset(lCount, 0) = cControl.MatchEntry
        .Names("Userforms.MouseIcon").RefersToRange.Offset(lCount, 0) = cControl.MouseIcon
        .Names("Userforms.MousePointer").RefersToRange.Offset(lCount, 0) = cControl.MousePointer
        .Names("Userforms.MultiSelect").RefersToRange.Offset(lCount, 0) = cControl.MultiSelect
        .Names("Userforms.Object").RefersToRange.Offset(lCount, 0) = cControl.Object
        .Names("Userforms.OldHeight").RefersToRange.Offset(lCount, 0) = cControl.OldHeight
        .Names("Userforms.OldLeft").RefersToRange.Offset(lCount, 0) = cControl.OldLeft
        .Names("Userforms.OldWidth").RefersToRange.Offset(lCount, 0) = cControl.OldWidth
        .Names("Userforms.Parent").RefersToRange.Offset(lCount, 0) = cControl.Parent
        .Names("Userforms.Picture").RefersToRange.Offset(lCount, 0) = cControl.Picture
        .Names("Userforms.PictureAlignment").RefersToRange.Offset(lCount, 0) = cControl.PictureAlignment
        .Names("Userforms.PicturePosition").RefersToRange.Offset(lCount, 0) = cControl.PicturePosition
        .Names("Userforms.PictureSizeMode").RefersToRange.Offset(lCount, 0) = cControl.PictureSizeMode
        .Names("Userforms.PictureTiling").RefersToRange.Offset(lCount, 0) = cControl.PictureTiling
        .Names("Userforms.RowSource").RefersToRange.Offset(lCount, 0) = cControl.RowSource
        .Names("Userforms.ScrollBars").RefersToRange.Offset(lCount, 0) = cControl.ScrollBars
        .Names("Userforms.ScrollHeight").RefersToRange.Offset(lCount, 0) = cControl.ScrollHeight
        .Names("Userforms.ScrollLeft").RefersToRange.Offset(lCount, 0) = cControl.ScrollLeft
        .Names("Userforms.ScrollTop").RefersToRange.Offset(lCount, 0) = cControl.ScrollTop
        .Names("Userforms.ScrollWidth").RefersToRange.Offset(lCount, 0) = cControl.ScrollWidth
        .Names("Userforms.Selected").RefersToRange.Offset(lCount, 0) = cControl.Selected
        .Names("Userforms.SpecialEffect").RefersToRange.Offset(lCount, 0) = cControl.SpecialEffect
        .Names("Userforms.TabIndex").RefersToRange.Offset(lCount, 0) = cControl.TabIndex
        .Names("Userforms.TabStop").RefersToRange.Offset(lCount, 0) = cControl.TabStop
        .Names("Userforms.Tag").RefersToRange.Offset(lCount, 0) = cControl.Tag
        .Names("Userforms.TakeFocusOnClick").RefersToRange.Offset(lCount, 0) = cControl.TakeFocusOnClick
        .Names("Userforms.Text").RefersToRange.Offset(lCount, 0) = cControl.Text
        .Names("Userforms.TextAlign").RefersToRange.Offset(lCount, 0) = cControl.TextAlign
        .Names("Userforms.TextColumn").RefersToRange.Offset(lCount, 0) = cControl.TextColumn
        .Names("Userforms.Title").RefersToRange.Offset(lCount, 0) = cControl.Title
        .Names("Userforms.Top").RefersToRange.Offset(lCount, 0) = cControl.Top
        .Names("Userforms.TopIndex").RefersToRange.Offset(lCount, 0) = cControl.TopIndex
        .Names("Userforms.TripleState").RefersToRange.Offset(lCount, 0) = cControl.TripleState
        .Names("Userforms.Value").RefersToRange.Offset(lCount, 0) = cControl.Value
        .Names("Userforms.VerticalScrollbarSide").RefersToRange.Offset(lCount, 0) = cControl.VerticalScrollBarSide
        .Names("Userforms.Visible").RefersToRange.Offset(lCount, 0) = cControl.Visible
        oSh.Names("Userforms.Width").RefersToRange.Offset(lCount, 0) = cControl.Width
        If cControl.Type = vbext_ct_MSForm Then
            .Names("Userforms.Code").RefersToRange.Offset(lCount, 0) = cControl.CodeModule.Lines(1, cControl.CodeModule.CountOfLines)
        End If
    End With
End Sub

Sub CreateForms()
    Dim oSh As Worksheet
    Dim ovbComp As VBComponent
    Dim oCtl As control
    Dim oForm As Object
    Dim lCount As Long
    Dim sFormName As String
    Dim sControlName As String
    Set oSh = ActiveSheet
    For lCount = 2 To Application.CountA(oSh.UsedRange.Columns(1).Cells.Value)
        If sFormName = oSh.Cells(lCount, 1) Then
            Set oCtl = oForm.Designer.Controls.Add("Forms." & oSh.Names("Userforms.Type").RefersToRange.Offset(lCount - 2, 0) & ".1", "temp", True)
            ChangeProperties sFormName, oCtl, oSh, lCount - 2
        Else
            sFormName = oSh.Cells(lCount, 1)
            Set oForm = ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
            ChangeProperties sFormName, oForm, oSh, lCount - 2
            oForm.Name = sFormName
            Set ovbComp = ActiveWorkbook.VBProject.VBComponents(sFormName)
            ovbComp.CodeModule.DeleteLines 1, ovbComp.CodeModule.CountOfLines
            ovbComp.CodeModule.AddFromString oSh.Names("Userforms.Code").RefersToRange.Offset(lCount - 2, 0)
            ovbComp.Properties("Width") = oSh.Names("Userforms.Width").RefersToRange.Offset(lCount - 2, 0)
            ovbComp.Properties("Height") = oSh.Names("Userforms.Height").RefersToRange.Offset(lCount - 2, 0)
        End If
    Next
End Sub

Sub ChangeProperties(ByVal sFormName As String, ByRef cControl As Object, ByRef oSh As Worksheet, ByVal lCount As Long)
    Dim oStartCell As Range
    On Error Resume Next
    Set oStartCell = oSh.Cells(lCount, 1)
    With cControl
        .Name = oSh.Names("Userforms.Controlname").RefersToRange.Offset(lCount, 0)
        If .Name = oStartCell.Offset(, 2) Then
            oSh.Names("Userforms.NewName").RefersToRange.Offset(lCount, 0) = .Name
        End If
        '    TypeName(cControl) = oSh.Names("Userforms.Type").RefersToRange.Offset(lCount, 0)
        .Accelerator = oSh.Names("Userforms.Accelerator").RefersToRange.Offset(lCount, 0)
        '    .ActiveControl = oSh.Names("Userforms.ActiveControl").RefersToRange.Offset(lCount, 0)
        .Alignment = oSh.Names("Userforms.Alignment").RefersToRange.Offset(lCount, 0)
        .AutoSize = oSh.Names("Userforms.AutoSize").RefersToRange.Offset(lCount, 0)
        .BackColor = oSh.Names("Userforms.BackColor").RefersToRange.Offset(lCount, 0)
        .BackStyle = oSh.Names("Userforms.BackStyle").RefersToRange.Offset(lCount, 0)
        .BorderColor = oSh.Names("Userforms.BorderColor").RefersToRange.Offset(lCount, 0)
        .BorderStyle = oSh.Names("Userforms.BorderStyle").RefersToRange.Offset(lCount, 0)
        .BoundColumn = oSh.Names("Userforms.BoundColumn").RefersToRange.Offset(lCount, 0)
        .BoundValue = oSh.Names("Userforms.BoundValue").RefersToRange.Offset(lCount, 0)
        .Cancel = oSh.Names("Userforms.Cancel").RefersToRange.Offset(lCount, 0)
        .CanPaste = oSh.Names("Userforms.CanPaste").RefersToRange.Offset(lCount, 0)
        .CanRedo = oSh.Names("Userforms.CanRedo").RefersToRange.Offset(lCount, 0)
        .CanUndo = oSh.Names("Userforms.CanUndo").RefersToRange.Offset(lCount, 0)
        .Caption = oSh.Names("Userforms.Caption").RefersToRange.Offset(lCount, 0)
        .Column = oSh.Names("Userforms.Column").RefersToRange.Offset(lCount, 0)
        .ColumnCount = oSh.Names("Userforms.ColumnCount").RefersToRange.Offset(lCount, 0)
        .ColumnHeads = oSh.Names("Userforms.ColumnHeads").RefersToRange.Offset(lCount, 0)
        .ColumnWidths = oSh.Names("Userforms.ColumnWidths").RefersToRange.Offset(lCount, 0)
        .ControlSource = oSh.Names("Userforms.ControlSource").RefersToRange.Offset(lCount, 0)
        .ControlTipText = oSh.Names("Userforms.ControlTipText").RefersToRange.Offset(lCount, 0)
        .Cycle = oSh.Names("Userforms.Cycle").RefersToRange.Offset(lCount, 0)
        .Default = oSh.Names("Userforms.Default").RefersToRange.Offset(lCount, 0)
        .DrawBuffer = oSh.Names("Userforms.DrawBuffer").RefersToRange.Offset(lCount, 0)
        .Enabled = oSh.Names("Userforms.Enabled").RefersToRange.Offset(lCount, 0)
        .Font = oSh.Names("Userforms.FontName").RefersToRange.Offset(lCount, 0)
        .ForeColor = oSh.Names("Userforms.FontSize").RefersToRange.Offset(lCount, 0)
        .ForeColor = oSh.Names("Userforms.ForeColor").RefersToRange.Offset(lCount, 0)
        .GroupName = oSh.Names("Userforms.GroupName").RefersToRange.Offset(lCount, 0)
        .Height = oSh.Names("Userforms.Height").RefersToRange.Offset(lCount, 0)
        .HelpContextID = oSh.Names("Userforms.HelpContextID").RefersToRange.Offset(lCount, 0)
        .IMEMode = oSh.Names("Userforms.IMEMode").RefersToRange.Offset(lCount, 0)
        .InsideHeight = oSh.Names("Userforms.InsideHeight").RefersToRange.Offset(lCount, 0)
        .InsideWidth = oSh.Names("Userforms.InsideWidth").RefersToRange.Offset(lCount, 0)
        .IntegralHeight = oSh.Names("Userforms.IntegralHeight").RefersToRange.Offset(lCount, 0)
        .KeepScrollBarsVisible = oSh.Names("Userforms.KeepScrollBarsVisible").RefersToRange.Offset(lCount, 0)
        '    .LayoutEffect = oSh.Names("Userforms.LayoutEffect").RefersToRange.Offset(lCount, 0)
        .Left = oSh.Names("Userforms.Left").RefersToRange.Offset(lCount, 0)
        .List = oSh.Names("Userforms.List").RefersToRange.Offset(lCount, 0)
        .ListCount = oSh.Names("Userforms.ListCount").RefersToRange.Offset(lCount, 0)
        .ListIndex = oSh.Names("Userforms.ListIndex").RefersToRange.Offset(lCount, 0)
        .ListStyle = oSh.Names("Userforms.ListStyle").RefersToRange.Offset(lCount, 0)
        .Locked = oSh.Names("Userforms.Locked").RefersToRange.Offset(lCount, 0)
        .MatchEntry = oSh.Names("Userforms.MatchEntry").RefersToRange.Offset(lCount, 0)
        .MouseIcon = oSh.Names("Userforms.MouseIcon").RefersToRange.Offset(lCount, 0)
        .MousePointer = oSh.Names("Userforms.MousePointer").RefersToRange.Offset(lCount, 0)
        .MultiSelect = oSh.Names("Userforms.MultiSelect").RefersToRange.Offset(lCount, 0)
        .Object = oSh.Names("Userforms.Object").RefersToRange.Offset(lCount, 0)
        '    .OldHeight = oSh.Names("Userforms.OldHeight").RefersToRange.Offset(lCount, 0)
        '    .OldLeft = oSh.Names("Userforms.OldLeft").RefersToRange.Offset(lCount, 0)
        '    .OldWidth = oSh.Names("Userforms.OldWidth").RefersToRange.Offset(lCount, 0)
        .Parent = oSh.Names("Userforms.Parent").RefersToRange.Offset(lCount, 0)
        .Picture = oSh.Names("Userforms.Picture").RefersToRange.Offset(lCount, 0)
        .PictureAlignment = oSh.Names("Userforms.PictureAlignment").RefersToRange.Offset(lCount, 0)
        .PicturePosition = oSh.Names("Userforms.PicturePosition").RefersToRange.Offset(lCount, 0)
        .PictureSizeMode = oSh.Names("Userforms.PictureSizeMode").RefersToRange.Offset(lCount, 0)
        .PictureTiling = oSh.Names("Userforms.PictureTiling").RefersToRange.Offset(lCount, 0)
        .RowSource = oSh.Names("Userforms.RowSource").RefersToRange.Offset(lCount, 0)
        .ScrollBars = oSh.Names("Userforms.ScrollBars").RefersToRange.Offset(lCount, 0)
        .ScrollHeight = oSh.Names("Userforms.ScrollHeight").RefersToRange.Offset(lCount, 0)
        .ScrollLeft = oSh.Names("Userforms.ScrollLeft").RefersToRange.Offset(lCount, 0)
        .ScrollTop = oSh.Names("Userforms.ScrollTop").RefersToRange.Offset(lCount, 0)
        .ScrollWidth = oSh.Names("Userforms.ScrollWidth").RefersToRange.Offset(lCount, 0)
        .Selected = oSh.Names("Userforms.Selected").RefersToRange.Offset(lCount, 0)
        .SpecialEffect = oSh.Names("Userforms.SpecialEffect").RefersToRange.Offset(lCount, 0)
        .TabIndex = oSh.Names("Userforms.TabIndex").RefersToRange.Offset(lCount, 0)
        .TabStop = oSh.Names("Userforms.TabStop").RefersToRange.Offset(lCount, 0)
        .Tag = oSh.Names("Userforms.Tag").RefersToRange.Offset(lCount, 0)
        .TakeFocusOnClick = oSh.Names("Userforms.TakeFocusOnClick").RefersToRange.Offset(lCount, 0)
        .Text = oSh.Names("Userforms.Text").RefersToRange.Offset(lCount, 0)
        .TextAlign = oSh.Names("Userforms.TextAlign").RefersToRange.Offset(lCount, 0)
        .TextColumn = oSh.Names("Userforms.TextColumn").RefersToRange.Offset(lCount, 0)
        .Title = oSh.Names("Userforms.Title").RefersToRange.Offset(lCount, 0)
        .Top = oSh.Names("Userforms.Top").RefersToRange.Offset(lCount, 0)
        .TopIndex = oSh.Names("Userforms.TopIndex").RefersToRange.Offset(lCount, 0)
        .TripleState = oSh.Names("Userforms.TripleState").RefersToRange.Offset(lCount, 0)
        .Value = oSh.Names("Userforms.Value").RefersToRange.Offset(lCount, 0)
        .VerticalScrollBarSide = oSh.Names("Userforms.VerticalScrollbarSide").RefersToRange.Offset(lCount, 0)
        .Visible = oSh.Names("Userforms.Visible").RefersToRange.Offset(lCount, 0)
        .Width = oSh.Names("Userforms.Width").RefersToRange.Offset(lCount, 0)
    End With
End Sub

