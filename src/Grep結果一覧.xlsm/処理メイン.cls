VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "処理メイン"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objWBK As Workbook
Private objWSH_貼り付け元 As Worksheet
Private objWSH_貼り付け先 As Worksheet

Private Sub Class_Initialize()
    'オブジェクトセット
    Set objWBK = ThisWorkbook
    Set objWSH_貼り付け元 = objWBK.Worksheets("【Grep結果貼り付け】")
    'Set objWSH_貼り付け先 = objWBK.Worksheets("【GREP調査フォーマット】")
    Set objWSH_貼り付け先 = objWBK.ActiveSheet  '複数のフォーマットシートを作成する事を考慮。
End Sub

Private Sub Class_Terminate()
    Set l_obj検索結果行 = Nothing
    Set objWSH_貼り付け元 = Nothing
    Set objWSH_貼り付け先 = Nothing
    Set objWBK = Nothing
End Sub

Public Sub make()

    Dim l_str検索文字列 As String
    Dim l_str検索結果1行 As String
    Dim l_obj検索結果行 As 検索結果行
    Dim l_lngCount As Long
    '2013/4/4 Ver.1.4追加 start
    Dim l_lngNo As Long
    Dim l_lngFirstRowCount As Long
    '2013/4/4 Ver.1.4追加 end
    
    '検索文字列を取得
    l_str検索文字列 = objWSH_貼り付け元.Cells(2, 1).Value
    l_str検索文字列 = Mid(l_str検索文字列, 8)
    l_str検索文字列 = Replace(l_str検索文字列, """", "")
    
    
    '検索最終行数を取得
    Dim l_lng検索最終行 As Long
    l_lng検索最終行 = objWSH_貼り付け元.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    '検索結果件数を取得
    Dim l_lng検索結果件数 As Long
    l_lng検索結果件数 = l_lng検索最終行 - 11
    
    
    'メイン処理（貼り付け元からデータ取得して、フォーマットへ転記する）
    For l_lngCount = 0 To l_lng検索結果件数 - 1
    
        l_str検索結果1行 = objWSH_貼り付け元.Cells(11 + l_lngCount, 1).Value
        '2013/4/4 Ver.1.4追加 start
        If l_str検索結果1行 = "" Then
            '検索結果行を取得できない場合は、なにも処理せずに次のレコードを取得する。
            '（前提として、検索結果の最初の行が取得できないケースを想定。
            '  検索結果の途中で行が取得できないケースはあり得ない想定としている。）
            l_lngFirstRowCount = l_lngFirstRowCount + 1
        Else
        '2013/4/4 Ver.1.4追加 end
    
            Set l_obj検索結果行 = New 検索結果行    'とりあえずパフォーマンスは無視。
            
            '＊＊No＊＊
            '2013/4/4 Ver.1.4修正 start
            'l_obj検索結果行.No = l_lngCount + 1
            l_lngNo = l_lngNo + 1
            l_obj検索結果行.No = l_lngNo
             '2013/4/4 Ver.1.4修正 start

            '＊＊検索文字列＊＊
            l_obj検索結果行.検索文字列 = l_str検索文字列
            
            '文字列切出しが必要な項目について処理
            l_obj検索結果行.文字列切出し_データセット (l_str検索結果1行)
            
         
        '2013/4/4 Ver.1.4追加 start
            If l_blFirstFlg = True Then
            End If
        '2013/4/4 Ver.1.4追加 start

            'フォーマットシートに転記
            With objWSH_貼り付け先
            
                '罫線作成
                '2013/4/4 Ver.1.修正 start
                'With .Range("A4:J4").Offset(l_lngCount, 0)
                With .Range("A4:J4").Offset(l_lngCount - l_lngFirstRowCount, 0)
                '2013/4/4 Ver.1.4修正 start
                
                        '上段
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
            
                    '下段
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
            
                    '左側
                    With .Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
            
                    '右側
                    With .Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
            
                    '内側の縦線
                    With .Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                End With
                
                '値入力
                '2013/4/4 Ver.1.修正 start
                '.Cells(4 + l_lngCount, 1).Value = l_obj検索結果行.No
                '.Cells(4 + l_lngCount, 2).Value = l_obj検索結果行.検索文字列
                '.Cells(4 + l_lngCount, 3).Value = l_obj検索結果行.フルパス
                '.Cells(4 + l_lngCount, 4).Value = l_obj検索結果行.ファイル名
                '.Cells(4 + l_lngCount, 5).Value = l_obj検索結果行.行数
                '.Cells(4 + l_lngCount, 7).Value = l_obj検索結果行.該当箇所

                .Cells(4 + l_lngCount - l_lngFirstRowCount, 1).Value = l_obj検索結果行.No
                .Cells(4 + l_lngCount - l_lngFirstRowCount, 2).Value = l_obj検索結果行.検索文字列
                .Cells(4 + l_lngCount - l_lngFirstRowCount, 3).Value = l_obj検索結果行.フルパス
                .Cells(4 + l_lngCount - l_lngFirstRowCount, 4).Value = l_obj検索結果行.ファイル名
                .Cells(4 + l_lngCount - l_lngFirstRowCount, 5).Value = l_obj検索結果行.行数
                .Cells(4 + l_lngCount - l_lngFirstRowCount, 7).Value = l_obj検索結果行.該当箇所
                '2013/4/4 Ver.1.修正 end

            End With
            
        '2013/4/4 Ver.1.4追加 start
        End If
        '2013/4/4 Ver.1.4追加 end
        
    Next l_lngCount



End Sub



