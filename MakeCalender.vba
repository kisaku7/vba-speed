Sub CreateYearlyCalendar()
    Dim yearInput As String
    Dim calendarYear As Integer
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim scheduleWs As Worksheet
    Dim hasSchedule As Boolean
    Dim scheduleDict As Object
    Dim lastUsedRow As Long
    
    ' 年の入力を受け取る
    yearInput = InputBox("カレンダーを作成する年を入力してください。", "カレンダー作成")
    If Not IsNumeric(yearInput) Or Len(yearInput) <> 4 Then
        MsgBox "正しい年を入力してください。"
        Exit Sub
    End If
    calendarYear = CInt(yearInput)
    
    hasSchedule = False
    Set scheduleDict = CreateObject("Scripting.Dictionary")
    
    ' スケジュールシートが存在するか確認
    On Error Resume Next
    Set scheduleWs = ThisWorkbook.Sheets("SCHEDULE")
    On Error GoTo 0
    If Not scheduleWs Is Nothing Then
        hasSchedule = True
        Dim lastRow As Long
        lastRow = scheduleWs.Cells(scheduleWs.Rows.Count, "A").End(xlUp).Row
        Dim i As Long
        For i = 1 To lastRow
            scheduleDict(CStr(scheduleWs.Cells(i, 1).Value)) = scheduleWs.Cells(i, 2).Value
        Next i
    End If
    
    ' 新しいブックを作成
    Set wb = Workbooks.Add
    
    Dim monthNames As Variant
    monthNames = Array("1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月")
    
    Dim startRow As Integer, startCol As Integer
    Dim firstDay As Date, lastDay As Date
    Dim dayOfWeek As Integer, dayCount As Integer
    
    For iMonth = 1 To 12
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = monthNames(iMonth - 1)
        
        ' 用紙設定
        With ws.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = 100
            .TopMargin = Application.CentimetersToPoints(0.5)
            .LeftMargin = Application.CentimetersToPoints(0.5)
            .BottomMargin = Application.CentimetersToPoints(0)
            .RightMargin = Application.CentimetersToPoints(0)
        End With
        
        ' 月の最初の日と最後の日を取得
        firstDay = DateSerial(calendarYear, iMonth, 1)
        lastDay = DateSerial(calendarYear, iMonth + 1, 0)
        
        ' カレンダーの見出しを作成
        ws.Cells(1, 1).Value = calendarYear & "年 " & iMonth & "月 (" & GetWareki(calendarYear) & ")"
        ws.Cells(1, 1).Font.Name = "メイリオ"
        ws.Cells(1, 1).Font.Size = 11
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).HorizontalAlignment = xlCenter
        ws.Cells(1, 1).VerticalAlignment = xlCenter
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 7)).Merge
        
        ' 曜日の見出しを作成
        For i = 0 To 6
            ws.Cells(2, i + 1).Value = WeekdayName(i + 1, True) & "曜日"
            ws.Cells(2, i + 1).Font.Name = "メイリオ"
            ws.Cells(2, i + 1).Font.Size = 11
            ws.Cells(2, i + 1).Font.Bold = True
            ws.Cells(2, i + 1).HorizontalAlignment = xlCenter
            ws.Cells(2, i + 1).VerticalAlignment = xlCenter
            If i = 0 Then
                ws.Cells(2, i + 1).Font.Color = RGB(255, 0, 0)
            ElseIf i = 6 Then
                ws.Cells(2, i + 1).Font.Color = RGB(0, 0, 255)
            End If
        Next i
        
        ' カレンダーの初日位置を決定
        startRow = 3
        startCol = Weekday(firstDay, vbSunday)
        
        ' 前月の日付を表示
        If startCol > 1 Then
            Dim prevMonthLastDay As Date
            prevMonthLastDay = DateSerial(calendarYear, iMonth, 0)
            For i = startCol - 1 To 1 Step -1
                ws.Cells(startRow, i).Value = Day(prevMonthLastDay)
                ws.Cells(startRow, i).Font.Color = RGB(128, 128, 128)
                ws.Cells(startRow, i).Font.Name = "メイリオ"
                ws.Cells(startRow, i).Font.Size = 11
                ws.Cells(startRow, i).HorizontalAlignment = xlLeft
                ws.Cells(startRow, i).VerticalAlignment = xlTop
                prevMonthLastDay = prevMonthLastDay - 1
            Next i
        End If
        
        ' 今月の日付を表示
        dayCount = 1
        Do While dayCount <= Day(lastDay)
            ws.Cells(startRow, startCol).Value = dayCount
            ws.Cells(startRow, startCol).Font.Name = "メイリオ"
            ws.Cells(startRow, startCol).Font.Size = 11
            ws.Cells(startRow, startCol).Font.Bold = True
            ws.Cells(startRow, startCol).HorizontalAlignment = xlLeft
            ws.Cells(startRow, startCol).VerticalAlignment = xlTop
            dayOfWeek = Weekday(DateSerial(calendarYear, iMonth, dayCount), vbSunday)
            If dayOfWeek = vbSunday Then
                ws.Cells(startRow, startCol).Font.Color = RGB(255, 0, 0)
            ElseIf dayOfWeek = vbSaturday Then
                ws.Cells(startRow, startCol).Font.Color = RGB(0, 0, 255)
            End If
            
            ' スケジュールの表示
            If hasSchedule And scheduleDict.exists(CStr(DateSerial(calendarYear, iMonth, dayCount))) Then
                ws.Cells(startRow + 1, startCol).Value = scheduleDict(CStr(DateSerial(calendarYear, iMonth, dayCount)))
                ws.Cells(startRow + 1, startCol).Font.Name = "メイリオ"
                ws.Cells(startRow + 1, startCol).Font.Size = 6
                ws.Cells(startRow + 1, startCol).HorizontalAlignment = xlLeft
                ws.Cells(startRow + 1, startCol).VerticalAlignment = xlTop
                ' 日曜日と土曜日のスケジュール部分のフォントカラーを設定
                If dayOfWeek = vbSunday Then
                    ws.Cells(startRow + 1, startCol).Font.Color = RGB(255, 0, 0)
                ElseIf dayOfWeek = vbSaturday Then
                    ws.Cells(startRow + 1, startCol).Font.Color = RGB(0, 0, 255)
                End If
            End If
            
            startCol = startCol + 1
            If startCol > 7 Then
                startCol = 1
                startRow = startRow + 2
            End If
            dayCount = dayCount + 1
        Loop
        
        ' 翌月の日付を表示（startColが1の場合は表示しない）
        If startCol > 1 And startCol <= 7 Then
            dayCount = 1
            Do While startCol <= 7
                ws.Cells(startRow, startCol).Value = dayCount
                ws.Cells(startRow, startCol).Font.Color = RGB(128, 128, 128)
                ws.Cells(startRow, startCol).Font.Name = "メイリオ"
                ws.Cells(startRow, startCol).Font.Size = 11
                ws.Cells(startRow, startCol).HorizontalAlignment = xlLeft
                ws.Cells(startRow, startCol).VerticalAlignment = xlTop
                startCol = startCol + 1
                dayCount = dayCount + 1
            Loop
        End If
        
        ' セルのサイズを設定
        ws.Columns("A:G").ColumnWidth = 8
        ws.Rows("1:2").RowHeight = 18.75
        
        ' 最後の行を取得
        lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        
        ' 行の高さを設定
        For row = 3 To lastUsedRow Step 2
            ws.Rows(row).RowHeight = 15
            ws.Rows(row + 1).RowHeight = 22
        Next row
        
        ' 枠線を追加（罫線）
        With ws.Range(ws.Cells(2, 1), ws.Cells(2, 7)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastUsedRow, 7)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastUsedRow, 7)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastUsedRow, 7)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastUsedRow, 7)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastUsedRow, 7)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastUsedRow, 7)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        ' 最後のスケジュール行に罫線を追加
        For col = 1 To 7
            ws.Cells(lastUsedRow, col).Borders.LineStyle = xlContinuous
            ws.Cells(lastUsedRow, col).Borders.Weight = xlThin
        Next col
        
        ' 日の行とスケジュール行の間は罫線を引かない
        For Row = 3 To lastUsedRow Step 2
            With ws.Range(ws.Cells(Row + 1, 1), ws.Cells(Row + 1, 7)).Borders(xlEdgeTop)
                .LineStyle = xlNone
            End With
        Next Row
        
        ' スケジュール部分を左上寄せの6ポイントに設定
        For row = 4 To lastUsedRow Step 2
            For col = 1 To 7
                With ws.Cells(row, col)
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlTop
                    .Font.Size = 6
                    .Font.Name = "メイリオ"
                End With
            Next col
        Next row
        
    Next iMonth
    
    ' "Sheet1" を削除
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Sheet1").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    MsgBox "カレンダーが作成されました。", vbInformation
End Sub

Function GetWareki(year As Integer) As String
    If year >= 2019 Then
        GetWareki = "令和" & (year - 2018)
    ElseIf year >= 1989 Then
        GetWareki = "平成" & (year - 1988)
    ElseIf year >= 1926 Then
        GetWareki = "昭和" & (year - 1925)
    ElseIf year >= 1912 Then
        GetWareki = "大正" & (year - 1911)
    Else
        GetWareki = "明治" & (year - 1867)
    End If
    GetWareki = GetWareki & "年"
End Function
