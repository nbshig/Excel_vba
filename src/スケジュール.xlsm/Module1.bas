Attribute VB_Name = "Module1"
Option Explicit

Sub Copy_Task(ByVal p_lngRowCount As Long, ByVal p_targetValue As String)

    Dim l_lngLastRowBeforTarget As Long
    Dim l_lngLastRowB As Long

    
    'A列のターゲット行より直前のA列日付の入力行を取得
    l_lngLastRowBeforTarget = Cells(p_lngRowCount, 1).End(xlUp).Row
    
    'B列上において、値が入力されている最終行を取得
    l_lngLastRowB = Cells(Rows.Count, 2).End(xlUp).Row

    '前週のタスクをコピー&ターゲット行にペースト
    Range(Cells(l_lngLastRowBeforTarget, 1), Cells(l_lngLastRowB, 1)).EntireRow.Copy _
        Destination:=Cells(p_lngRowCount, 1)
    
    'A列ターゲット行に値を入力
    Cells(p_lngRowCount, 1) = p_targetValue
    
End Sub


Sub Set_FormatConditions(ByVal p_lngNewLastRowB As Long)

    Dim fc As FormatCondition
    
    '条件付き書式を削除
    Cells.FormatConditions.Delete
        
    
    '=$I1="完了" →=B列1行目:I列最終行の範囲で薄い灰色
    With Range("I1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$I1=""完了""")
        fc.Interior.Color = XlRgbColor.rgbLightGray
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With


    '=$G1="完了" →=B列1行目:I列最終行の範囲で薄い灰色
    With Range("G1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$G1=""完了""")
        fc.Interior.Color = XlRgbColor.rgbLightGray
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With


    '=$B1="No" →=B列1行目:I列最終行の範囲で淡いアクアマリン色
    With Range("B1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$B1=""No""")
        fc.Interior.Color = XlRgbColor.rgbMediumAquamarine
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With


    '=AND($B1<>"No", $B1<>"",MOD(ROW(),2)=0) →=B列1行目:I列最終行の範囲でリネン色
    With Range("B1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($B1<>""No"", $B1<>"""",MOD(ROW(),2)=0)")
        fc.Interior.Color = XlRgbColor.rgbLinen
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With
End Sub
