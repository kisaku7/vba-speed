Option Explicit

' src: https://tonari-it.com/excel-vba-class-performance-booster/


Private initCalculationValue_ As XlCalculation

Private Sub Class_Initialize()

    With Application
        initCalculationValue_ = .Calculation
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

End Sub

Private Sub Class_Terminate()

    With Application
        .Calculation = initCalculationValue_
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub

