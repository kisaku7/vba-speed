Option Explicit

' 現在のシートに貼り付け後、別ブックにして保存
Sub moveToSheetAndSave(ByVal rng As Variant, ByVal saveTo As String)
  rng.Copy
  Dim moveSheet As Variant: Set moveSheet = rng.Parent.Parent.Worksheets.Add
  moveSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
  Dim moveBook As Variant: Set moveBook = Workbooks.Add
  moveSheet.Move before:=moveBook.Worksheets(1)
  Set moveSheet = moveBook.Worksheets(1)
  Application.DisplayAlerts = False
  moveSheet.Parent.SaveAs Filename:=saveTo, FileFormat:=xlCurrentPlatformText
  moveBook.Close saveChanges:=False
  Application.DisplayAlerts = True
End Sub
Sub t_moveToSheetAndSave()
  Dim path As Variant: path = getSavePath
  If path = False Then
    Debug.Print "未選択"
    Exit Sub
  End If
  
  Call moveToSheetAndSave(Range("A1:B5"), path)
End Sub



Sub moveToBookAndSave(ByVal rng As Variant, ByVal saveTo As String)
  rng.Copy
  Dim moveSheet As Variant: Set moveSheet = Workbooks.Add.Worksheets(1)
  moveSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
  Application.DisplayAlerts = False
  With moveSheet.Parent
    .SaveAs Filename:=saveTo, FileFormat:=xlCurrentPlatformText
    .Close saveChanges:=False
  End With
  Application.DisplayAlerts = True
End Sub
Sub t_moveToBookAndSave()
  Dim path As Variant: path = getSavePath
  If path = False Then
    Debug.Print "未選択"
    Exit Sub
  End If
  
  Call moveToBookAndSave(Range("A1:B5"), path)
End Sub


' moveToSheet
'  15000row*3: 18~21sec
' moveToBook
'  15000row: 6sec
Sub test()
  
  Dim path As Variant: path = getSavePath
  If path = False Then
    Debug.Print "未選択"
    Exit Sub
  End If
  
  Dim timerObj As TimerObject: Set timerObj = New TimerObject
  Dim booster As PerformanceBooster: Set booster = New PerformanceBooster
  Dim i As Long
  For i = 1 To 3
    'Call moveToSheetAndSave(Range("A1:GS15000"), path)
    'Call moveToBookAndSave(Range("A1:GS5000"), path)
    Call test4(Range("A1:GS15000"), 15000)
  Next
  timerObj.ReportTimer
End Sub





Function getSavePath() As Variant
  getSavePath = Application.GetSaveAsFilename(InitialFileName:="testoutput.txt", Filefilter:="Text Files (*.txt), *.txt")
End Function
Sub t_getSavePath()
  Dim path As Variant: path = getSavePath
  If path = False Then
    Debug.Print "未選択"
  Else
    Debug.Print "path:" & path
  End If
End Sub
