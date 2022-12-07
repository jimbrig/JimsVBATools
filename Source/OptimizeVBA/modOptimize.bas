' ------------------------------------------------------
' Name: modOptimize
' Kind: Module
' Purpose: VBA Optimization Utility Helpers
' Author: Jimmy Briggs
' Date: 2022-06-02
' ------------------------------------------------------
Option Explicit

' ------------------------------------------------------
' Section: Declarations
'-------------------------------------------------------

' Settings - Logging
' Public EnableLogging As Boolean
' Public Logger As cldLogger

' Set Previous State Public Dims
Dim prevCalc As XlCalculation       ' Application.Calculation
Dim prevEvents As Boolean           ' Application.EnableEvents
Dim prevScreen As Boolean           ' Application.ScreenUpdating
Dim prevAlerts As Boolean          ' Application.DisplayAlerts
Dim prevStatusBar As Boolean        ' Application.StatusBar
Dim prevDisplayStatusBar As Boolean ' Application.DisplayStatusBar
Dim prevAnimations As Boolean       ' Application.EnableAnimations
Dim prevPageBreaks As Boolean       ' ActiveSheet.DisplayPageBreaks

' ----------------------------------------------------------------
' Procedure Name: OptimizeVBA
' Purpose: Toggle VBA Optimizations Depending On Current State
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter isOn (Boolean): Ttru to Run OptimizeOn() and False to Run OptimizeOff()
' Author: Jimmy Briggs
' Date: 2022-06-02
' ----------------------------------------------------------------
Public Sub OptimizeVBA(ByVal Toggle As Boolean)
    
    ' EnableLogging = GetSetting("EableLogging")
    
    With Application
        .ScreenUpdating = Not Toggle
        .EnableEvents = Not Toggle
        .DisplayAlerts = Not Toggle
        .StatusBar = Not Toggle
        .EnableAnimations = Not Toggle
        .DisplayStatusBar = Not Toggle
        .Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
    End With
    
    ActiveSheet.DisplayPageBreaks = Not Toggle

    ' If Toggle And EnableLogging Then
    '   Set Logger = New clsLogger
    '   LogDebug "Start Logging..."
    ' End If
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: OptimizeOn
' Purpose: Toggle VBA Optimizations ON to Optimize VBA exectution
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Jimmy Briggs
' Date: 2022-06-02
' ----------------------------------------------------------------
Sub OptimizeOn()
    prevCalc = Application.Calculation: Application.Calculation = xlCalculationManual
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevPageBreaks = ActiveSheet.DisplayPageBreaks: ActiveSheet.DisplayPageBreaks = False
    prevStatusBar = Application.StatusBar: Application.StatusBar = False
    prevDisplayStatusBar = Application.DisplayStatusBar: Application.DisplayStatusBar = False
    prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    prevAnimations = Application.EnableAnimations: Application.EnableAnimations = False
End Sub

' ----------------------------------------------------------------
' Procedure Name: OptimizeOff
' Purpose: Toggle VBA Optimizations OFF. Turn off VBA Optimizations
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Jimmy Briggs
' Date: 2022-06-02
' ----------------------------------------------------------------
Sub OptimizeOff()
    With Application
        .Calculation = prevCalc
        .EnableEvents = prevEvents
        .ScreenUpdating = prevScreen
        .DisplayStatusBar = prevDisplayStatusBar
        .StatusBar = prevStatusBar
        .DisplayAlerts = prevAlerts
    End With
    ActiveSheet.DisplayPageBreaks = prevPageBreaks
End Sub
