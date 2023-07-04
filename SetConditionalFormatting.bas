Sub SetConditionalFormatting()
    Dim progress_area As Range, status_area As Range, status_allcolumn_area As Range, task_area As Range
    Dim ws As Worksheet
    
    ' 適切なワークシートを設定します（ワークシート名を必要に応じて変更してください）
    Set ws = ThisWorkbook.Sheets("WBS")
    
    Set progress_area = ws.Range("$M$5:$GJ$2000")
    Set status_area = ws.Range("$F$5:$F$2000")
    Set status_allcolumn_area = ws.Range("$B$5:$L$2000")
    Set task_area = ws.Range("$C$5:$E$2000")
    
    ' 既存の条件付き書式を削除します
    ws.Cells.FormatConditions.Delete
    
    ' 条件１
    With progress_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($K5<>"""",M$2>=$K5,OR(M$2<=$L5,AND($L5="""",M$2<=TODAY())))")
        .Interior.Color = RGB(30, 80, 181) ' 青色
    End With
    
    ' 条件２
    With progress_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($J5<>"""",M$2>=$I5,M$2<=$J5)")
        .Interior.Color = RGB(218, 227, 243) ' 灰青色
    End With
    
    ' 条件３
    With progress_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=IF(COUNTIF(holidays,M$4),TRUE,FALSE)")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件４
    With progress_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=WEEKDAY(M$4)=1")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件５
    With progress_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=WEEKDAY(M$4)=7")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件６
    With status_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""遅延"", $F5))")
        .Font.Color = RGB(255, 0, 0) ' 赤色
    End With
    
    ' 条件７
    With status_allcolumn_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""完了"", $F5))")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件８
    With status_allcolumn_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($I5<(TODAY()+15),$I5>(TODAY()-14),OR($F5=""未着手"",""開始遅延""))")
        .Interior.Color = RGB(255, 217, 102) ' オレンジ色
    End With

    ' 条件９
    With task_area.FormatConditions.Add(Type:=xlExpression, Formula1:="=C4=C5")
        .Font.Color = RGB(240, 240, 240) ' 白色
    End With
    
End Sub
