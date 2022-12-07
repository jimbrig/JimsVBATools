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

' Set Previous State Dims
Dim prevCalc As XlCalculation ' Application.Calculation
Dim prevEvents As Boolean     ' Application.EnableEvents
Dim prevScreen As Boolean     ' Application.ScreenUpdating
Dim prevPageBreaks As Boolean ' ActiveSheet.DisplayPageBreaks

' ----------------------------------------------------------------
' Procedure Name: OptimizeVBA
' Purpose: Toggle VBA Optimizations Depending On Current State
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter isOn (Boolean): Ttru to Run OptimizeOn() and False to Run OptimizeOff()
' Author: Jimmy Briggs
' Date: 2022-06-02
' ----------------------------------------------------------------
Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
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
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    ActiveSheet.DisplayPageBreaks = prevPageBreaks
End Sub
