Sub SetConditionalFormatting()
    Dim rng1 As Range, rng2 As Range, rng3 As Range
    Dim ws As Worksheet
    
    ' 適切なワークシートを設定します（ワークシート名を必要に応じて変更してください）
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Set rng1 = ws.Range("$M$5:$GJ$2000")
    Set rng2 = ws.Range("$F$5:$F$2000")
    Set rng3 = ws.Range("$C$5:$E$2000")
    
    ' 既存の条件付き書式を削除します
    ws.Cells.FormatConditions.Delete
    
    ' 条件１
    With rng1.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($K5<>"""",M$2>=$K5,OR(M$2<=$L5,AND($L5="""",M$2<=TODAY())))")
        .Interior.Color = RGB(0, 0, 255) ' 青色
    End With
    
    ' 条件２
    With rng1.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($J5<>"""",M$2>=$I5,M$2<=$J5)")
        .Interior.Color = RGB(128, 128, 192) ' 灰青色
    End With
    
    ' 条件３
    With rng1.FormatConditions.Add(Type:=xlExpression, Formula1:="=IF(COUNTIF(holidays,M$4),TRUE,FALSE)")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件４
    With rng1.FormatConditions.Add(Type:=xlExpression, Formula1:="=WEEKDAY(M$4)=1")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件５
    With rng1.FormatConditions.Add(Type:=xlExpression, Formula1:="=WEEKDAY(M$4)=7")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件６
    With rng2.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""遅延"", $F5))")
        .Font.Color = RGB(255, 0, 0) ' 赤色
    End With
    
    ' 条件７
    With rng2.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""完了"", $F5))")
        .Interior.Color = RGB(192, 192, 192) ' 灰色
    End With
    
    ' 条件８
    With rng3.FormatConditions.Add(Type:=xlExpression, Formula1:="=C4=C5")
        .Font.Color = RGB(255, 255, 255) ' 白色
    End With

End Sub