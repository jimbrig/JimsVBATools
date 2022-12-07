' ------------------------------------------------------------------
' Name: modOptimize
' Kind: Module
' Purpose: VBA Optimization Utility Helpers
' Author: Jimmy Briggs
' Date: 2022-06-02
' ------------------------------------------------------------------
Option Explicit

' ------------------------------------------------------------------
' Section: Documentation
' ------------------------------------------------------------------
' VBA Optimization Utility Sub Procedures and Best Practices
' ------------------------------------------------------------------
' Toggles:
'   Application.Calculation
'   Application.EnableEvents
'   Application.ScreenUpdating
'   Application.DisplayAlerts
'   Application.DisplayStatusBar and Application.StatusBar
'   Application.EnableAnimations
'   Application.PrintCommunication
'   ActiveSheet.DisplayPageBreaks
' ------------------------------------------------------------------
' The following Procedures are in this module:
'   OptimizeVBA: Toggle VBA Optimizations Depending On Current State
'   OptimizeOn: Toggle VBA Optimizations ON
'   OptimizeOff: Toggle VBA Optimizations OFF
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' Section: Declarations
' ------------------------------------------------------------------
' Set Previous State Public Dims
Public Dim prevCalc As XlCalculation           ' Application.Calculation
Public Dim prevEvents As Boolean               ' Application.EnableEvents
Public Dim prevScreen As Boolean               ' Application.ScreenUpdating
Public Dim prevAlerts As Boolean               ' Application.DisplayAlerts
Public Dim prevStatusBar As Boolean            ' Application.StatusBar
Public Dim prevDisplayStatusBar As Boolean     ' Application.DisplayStatusBar
Public Dim prevAnimations As Boolean           ' Application.EnableAnimations
Public Dim prevPageBreaks As Boolean           ' ActiveSheet.DisplayPageBreaks
Public Dim prevPrintCommunication As Boolean   ' Application.PrintCommunication

' Settings - Logging - Uncomment to Enable
' Public EnableLogging As Boolean
' Public Logger As clsLogger
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' Section: Procedures
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' OptimizeVBA
' ------------------------------------------------------------------
' Procedure Name: OptimizeVBA
' Purpose: Toggle VBA Optimizations Depending On Current State
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter isOn (Boolean): True to Run OptimizeOn() and False to Run OptimizeOff()
' Author: Jimmy Briggs
' Date: 2022-06-02
' Example: 
' ```vba
' Public Sub DoStuff()
'   OptimizeVBA True
'   <Do Stuff>
'   OptimizeVBA False
' End Sub
' ```
' -----------------------------------------------------------------
Public Sub OptimizeVBA(ByVal Toggle As Boolean)    
    ' EnableLogging = GetSetting("EableLogging")    
    With Application
        .ScreenUpdating = Not Toggle
        .EnableEvents = Not Toggle
        .DisplayAlerts = Not Toggle
        .EnableAnimations = Not Toggle
        .DisplayStatusBar = Not Toggle
        .StatusBar = Not Toggle
        .PrintCommunication = Not Toggle
        .Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
    End With    
    ActiveSheet.DisplayPageBreaks = Not Toggle
    ' If Toggle And EnableLogging Then
    '   Set Logger = New clsLogger
    '   LogDebug "Start Logging..."
    ' End If    
End Sub

' Public Sub OptimizeVBA2(ByVal isOn As Boolean)
'    If isOn Then
'        OptimizeOn
'    Else
'        OptimizeOff
'    End If
' End Sub

' ------------------------------------------------------------------
' OptimizeOn
' ------------------------------------------------------------------
' Procedure Name: OptimizeOn
' Purpose: Toggle VBA Optimizations ON to Optimize VBA exectution
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Jimmy Briggs
' Date: 2022-06-02
' ----------------------------------------------------------------
Public Sub OptimizeOn()
    With Application
        prevCalc = .Calculation: .Calculation = xlCalculationManual
        prevEvents = .EnableEvents: .EnableEvents = False
        prevScreen = .ScreenUpdating: .ScreenUpdating = False
        prevDisplayStatusBar = .DisplayStatusBar: .DisplayStatusBar = False
        prevStatusBar = .StatusBar: .StatusBar = False
        prevPrintCommunication = .PrintCommunication: .PrintCommunication = False
        prevAlerts = .DisplayAlerts: .DisplayAlerts = False
        prevAnimations = .EnableAnimations: .EnableAnimations = False
    End With
    With ActiveSheet
        prevPageBreaks = .DisplayPageBreaks: .DisplayPageBreaks = False
    End With
End Sub

' ------------------------------------------------------------------
' OptimizeOff
' ------------------------------------------------------------------
' Procedure Name: OptimizeOff
' Purpose: Toggle VBA Optimizations OFF. Turn off VBA Optimizations
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Jimmy Briggs
' Date: 2022-06-02
' ----------------------------------------------------------------
Public Sub OptimizeOff()
    With Application
        .Calculation = prevCalc
        .EnableEvents = prevEvents
        .ScreenUpdating = prevScreen
        .DisplayStatusBar = prevDisplayStatusBar
        .StatusBar = prevStatusBar
        .DisplayAlerts = prevAlerts
        .EnableAnimations = prevAnimations
        .PrintCommunication = prevPrintCommunication
    End With
    With ActiveSheet
        .DisplayPageBreaks = prevPageBreaks
    End With
End Sub

' ------------------------------------------------------------------
' Section: Notes
' ------------------------------------------------------------------
' Toggles Not Implemented / Ignored:
'   Application.DisplayFormulaBar
'   Application.DisplayNoteIndicator
'   Application.DisplayRecentFiles
'   Application.DisplayScrollBars
'   Application.DisplayWorkbookTabs
'   Application.DisplayXMLSourcePane
'   Application.EnableAutoComplete
'   Application.EnableCancelKey
'   Application.EnableSound
'   Application.EnableTipWizard
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' Section: References
' ------------------------------------------------------------------
'  https://vbacompiler.com/optimize-vba-code/
'  https://analysistabs.com/vba/optimize-code-run-macros-faster/
'  https://www.spreadsheet1.com/how-to-optimize-vba-performance.html
'  https://eident.co.uk/2016/03/top-ten-tips-to-speed-up-your-vba-code/
'  https://excelitems.com/2010/12/optimize-vba-code-for-faster-macros.html
' ------------------------------------------------------------------

