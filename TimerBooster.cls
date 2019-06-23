Option Explicit


' src: https://tonari-it.com/excel-vba-class-timer


Public Start As Date
Public Finish As Date
 
Private Sub Class_Initialize()
    Start = Time
End Sub
 
Public Sub ReportTimer()
    Finish = Time
    MsgBox "実行時間は " & Format(Finish - Start, "nn分ss秒") & " でした", vbInformation + vbOKOnly
End Sub

