Attribute VB_Name = "Module1"
Option Explicit
'################################
'メンバ変数宣言
'################################
Dim m_intStartColumn As Integer         '管理対象となる日毎管理の開始列
Dim m_intEndCoumn As Integer            '管理対象となる日毎管理の終了列
Dim m_lngStartRow As Long               '管理対象となるケースの開始行
Dim m_lngEndRow As Long                 '管理対象となるケースの終了行
Dim m_intStartColumnRuiseki As Integer  '累積項目明細の開始列
Dim m_intEndCoumnRuiseki As Integer     '累積項目明細の終了列
Public m_str呼出し元処理 As String

Sub 日毎予実管理追加()

    Application.ScreenUpdating = False  '画面更新の非表示

    '##############
    'メンバ変数初期化
    '##############
    m_intStartColumn = 18   '18列目(開始列)
    m_intEndCoumn = 0
    m_lngStartRow = 4       '4行目(開始行)
    m_lngEndRow = Cells(Rows.Count, 2).End(xlUp).Row    '最後の行
    m_intStartColumnRuiseki = 0
    m_intEndCoumnRuiseki = 0


    '################################
    '変数宣言
    '################################
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    ReDim l_arrHiduke(2) As Date
    Dim l_intNissu As Integer
    Dim l_lngWrkCellY As Long
    Dim l_lngWrkCellYWeek As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim shp As CheckBox
    Dim l_chkFlg As Boolean


    '共通チェック
    If 共通チェック = False Then
        Exit Sub
    End If


    '全体作成 or 部分作成の判断を行う
    On Error GoTo ErrorHandler_WoekSheetBKUP        'バックアップシートが残っているようならエラーメッセージを表示
    
    If TypeName(Selection) <> "Range" Then
        'セル以外が選択されているときに実行時エラーが発生しないように終了する
        MsgBox "セル以外が選択されている可能性があります。" & vbCrLf & "任意のセルを選択した状態にしておいて下さい。"
        Exit Sub
    Else
        For Each shp In ActiveSheet.CheckBoxes
            With shp
            If .TopLeftCell.Address = "$B$1" Then   '日次明細列の値削除を確認するチェックボックス
                If .Value = 1 Then                 'チェックボックスがONのとき
                    l_chkFlg = False
                    .Value = -4146                 'チェックボックスを初期化
                Else                               'チェックボックスがOFFのとき
                    l_chkFlg = True
                    'バックアップシート作成
                    Worksheets("日次管理").Copy after:=Sheets("日次管理")
                    ActiveSheet.name = "bk_日次管理"
                    Worksheets("日次管理").Activate
                    Exit For
                End If
            End If
            End With
        Next
    End If
    
    '##############
    'クリア処理
    '##############
    Call クリア処理
    
    
    '################################
    '最小の日付、最大の日付を取得する
    '################################
    l_arrHiduke = 日付取得()
    l_datHidukeMin = l_arrHiduke(0) '最小日付
    l_datHidukeMax = l_arrHiduke(1) '最大日付
    
    
    '最小日付と最大日付の間の日数を取得する
    l_intNissu = DateDiff("d", l_datHidukeMin, l_datHidukeMax)


    '######################################################################
    '######################################################################
    '日次明細列の作成（最小日付から最大日付までの日数分）
    '######################################################################
    '######################################################################
    
    '##############
    'タイトル設定
    '##############
    l_lngWrkCellY = タイトル設定(l_intNissu, l_datHidukeMin)


    '##############
    'メンバ変数セット
    '##############
    '管理対象となる日毎管理の終了列をセット
    m_intEndCoumn = l_lngWrkCellY - 1
    
    '累積項目明細の開始列をセット
    m_intStartColumnRuiseki = m_intEndCoumn + 1
    
    '累積項目明細の終了列をセット
    m_intEndCoumnRuiseki = m_intEndCoumn + 4
    
    
    '##############
    '累積項目を設定する
    '##############
    Call 累積項目追加
    
    
    '###########################
    '日次合計/累積項目を設定する
    '###########################
    Call 日次合計_累積項目追加



    '##############
    '罫線描写
    '##############
    Call 罫線描写(l_intNissu)


    '##############
    '表示形式設定
    '##############
    Call 表示形式設定


    '##############
    'フォーマット用背景色設定
    '##############
    '一時保管用変数
    Dim l_lngWkSR As Long
    Dim l_lngWkSC As Long
    Dim l_lngWkER As Long
    Dim l_lngWkEC As Long
    
    l_lngWkSR = m_lngStartRow
    l_lngWkSC = m_intStartColumn
    l_lngWkER = m_lngEndRow
    l_lngWkEC = m_intEndCoumn
    
    m_lngStartRow = 4
    m_intStartColumn = 1
    m_intEndCoumn = 17
    
    
    Call フォーマット用背景色設定
    
    '元のメンバ変数の値を戻す
    m_lngStartRow = l_lngWkSR
    m_intStartColumn = l_lngWkSC
    m_intEndCoumn = l_lngWkEC


    '##############
    '背景色設定
    '##############
    Call 背景色設定(l_intNissu)


    '###########################
    'チェックボックス作成
    '###########################
    Call CheckBox作成


    '###########################
    'チェックがある場合、ここでバックアップで取った値を戻す
    '###########################
    If l_chkFlg Then
        Call 値再設定(l_datHidukeMin, l_datHidukeMax)
    End If


    '###########################
    '実績日を登録する
    '###########################
    Call 実績日登録
    
    
    '###########################
    'バックアップシートの削除
    '###########################
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.name = "bk_日次管理" Then
            Application.DisplayAlerts = False
            Worksheets("bk_日次管理").Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws


    '######################################################################
    '######################################################################
    '予実管理シートの作成
    '######################################################################
    '######################################################################
    Call 予実管理作成(l_intNissu)
    
    '19/04/04 Add Start
    '######################################################################
    '######################################################################
    '予実管理(機能単位)シートの作成
    '######################################################################
    '######################################################################
    Call 予実管理作成_機能単位(l_intNissu)
    '19/04/04 Add End
    
    
    MsgBox "作成完了"

    Application.ScreenUpdating = True  '画面更新の表示
    
    Exit Sub

    '###########################
    '例外処理
    '###########################
ErrorHandler_WoekSheetBKUP:
    MsgBox Err.Number & ":" & Err.Description & vbCrLf & "(又は、bk_日次管理シートが残っている可能性があります。削除して下さい。)", vbCritical & vbOKOnly, "エラー"

End Sub

Private Function 日付取得() As Date()

    '################################
    '変数宣言
    '################################
    Dim l_datHidukeArry(2) As Date
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_datwkHiduke  As Date
    Dim i As Integer


    '################################
    '変数初期化
    '################################
    l_datHidukeMin = "9999/12/31"
    l_datHidukeMax = "1900/01/01"
    
    
    '################################
    '(テスト実施) 着手予定日/着手実績日の中で最小の日付を取得する
    '################################
    '(テスト実施) 最小の着手予定日を取得する
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 7) = "" Then
     Else
      l_datwkHiduke = Cells(i, 7).Value
      If l_datwkHiduke < l_datHidukeMin Then
        l_datHidukeMin = l_datwkHiduke
      End If
     End If
    Next i
    
    '(テスト実施) 最小の着手実績日を取得する
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 8) = "" Then
     Else
      l_datwkHiduke = Cells(i, 8).Value
      If l_datwkHiduke < l_datHidukeMin Then
        l_datHidukeMin = l_datwkHiduke
      End If
     End If
    Next i
    
    '################################
    '(テスト実施) 取得した最小の日付から直近が月曜日となる日付を取得する
    '################################
    l_datHidukeMin = searchMonday(l_datHidukeMin)
    
    
    '################################
    '(テスト実施) 完了予定日/完了実績日を取得する、及び
    '(テスト精査) 完了予定日/完了実績日の中で、最大の日付を取得する
    '################################
    '(テスト実施) 最大の完了予定日を取得する
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 9) = "" Then
     Else
      l_datwkHiduke = Cells(i, 9).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i
    
    '(テスト実施) 最大の完了実績日を取得する
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 10) = "" Then
     Else
      l_datwkHiduke = Cells(i, 10).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i
    
    
    '(テスト精査) 最大の完了予定日を取得する
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 12) = "" Then
     Else
      l_datwkHiduke = Cells(i, 12).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i
    
    '(テスト精査) 最大の完了実績日を取得する
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 13) = "" Then
     Else
      l_datwkHiduke = Cells(i, 13).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i


    '################################
    '(テスト実施) 取得した最大の日付から直後が日曜日となる日付を取得する
    '################################
    l_datHidukeMax = searchSunday(l_datHidukeMax)
    
    
    '################################
    '最小日付、最大日付を返す
    '################################
    l_datHidukeArry(0) = l_datHidukeMin
    l_datHidukeArry(1) = l_datHidukeMax

    日付取得 = l_datHidukeArry
    
End Function


'################################
'引数の日付から直近が月曜日となる日付を返す
'################################
Private Function searchMonday(ByVal p_datTmpDate As Date) As Date

    Do While Weekday(p_datTmpDate) <> vbMonday
        p_datTmpDate = p_datTmpDate - 1
    Loop

    searchMonday = p_datTmpDate
    
End Function


'################################
'引数の日付から直後が日曜日となる日付を返す
'################################
Private Function searchSunday(ByVal p_datTmpDate As Date) As Date

    Do While Weekday(p_datTmpDate) <> vbSunday
        p_datTmpDate = p_datTmpDate + 1
    Loop

    searchSunday = p_datTmpDate
    
End Function


'################################
'累積項目を設定する
'################################
Private Sub 累積項目追加()

    Dim l_lngWrkCellY As Long
    Dim j As Integer
    Dim i As Integer
    Dim l_rngAddress As Range


    '############################
    'タイトル値と各累積値を入力
    '############################
    '変数初期化
    l_lngWrkCellY = 1
    
    For j = m_intStartColumnRuiseki To m_intEndCoumnRuiseki
            If (j Mod 4) = 2 Then                               '実施完了-予定
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "累積"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "実施完了"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "予定"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=2)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            ElseIf (j Mod 4) = 3 Then                           '実施完了-実績
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "累積"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "実施完了"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "実績"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=3)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            ElseIf (j Mod 4) = 0 Then                           '精査完了-予定
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "累積"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "精査完了"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "予定"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=0)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            ElseIf (j Mod 4) = 1 Then                           '精査完了-実績
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "累積"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "精査完了"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "実績"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=1)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            Else
                '処理なし
            End If
            
            l_lngWrkCellY = l_lngWrkCellY + 1
    Next j

End Sub

'################################
'日次合計_累積項目を設定する
'################################
Private Sub 日次合計_累積項目追加()

    Dim l_lngWrkCellY As Long
    Dim l_rngAddress As Range
    Dim j As Integer
    
    For j = m_intStartColumn To m_intEndCoumnRuiseki
        '日次合計の計算式を設定
        Set l_rngAddress = Range(Cells(m_lngStartRow, j), Cells(m_lngEndRow, j))
        Cells(m_lngEndRow + 1, j) = "=SUM(" + l_rngAddress.Address + ")"
        
        '累積の計算式を設定
        If j <= m_intStartColumn + 3 Then                       '最初の日にちのケース
            Cells(m_lngEndRow + 2, j) = "=" + Cells(m_lngEndRow + 1, j).Address
        ElseIf j <= m_intEndCoumn Then                          '最初の日にちより後、累積項目までのケース
            Cells(m_lngEndRow + 2, j) = "=" + Cells(m_lngEndRow + 1, j).Address + "+" + Cells(m_lngEndRow + 2, j - 4).Address
        Else
            '処理なし（累積項目を想定）
        End If
        
    Next j

End Sub


'################################
'クリア処理
'################################
Public Sub クリア処理()
    Dim shp As CheckBox

    '日次明細列の削除
    Range("R1", Cells(1, Columns.Count)).EntireColumn.Delete
    
    '実績日の入力値削除
    Range(Cells(m_lngStartRow, 8), Cells(m_lngEndRow, 8)).Value = ""        'テスト実施-着手-実績
    Range(Cells(m_lngStartRow, 10), Cells(m_lngEndRow, 10)).Value = ""      'テスト実施-完了-実績
    Range(Cells(m_lngStartRow, 13), Cells(m_lngEndRow, 13)).Value = ""      'テスト精査-完了-実績
    
    '全てのチェックボックスを削除する
    For Each shp In ActiveSheet.CheckBoxes
        If shp.TopLeftCell.Address <> "$B$1" Then
                shp.Delete                                                  '日次明細列の値削除を確認するチェックボックス以外は削除する
        End If
    Next


End Sub


'################################
'チェックボックス作成
'################################
Sub CheckBox作成()

    Dim i As Integer

    'C列に値がある行についてA列にチェックボックスを作成する
    For i = m_lngStartRow To m_lngEndRow
        Cells(i, 1).Activate
        
        With ActiveSheet.CheckBoxes.Add(1, 1, 1, 1)
            .Height = 10
            .Top = ActiveCell.Top
            .Left = ActiveCell.Left
            .Caption = ""                           'テキスト
            .Value = False
        End With
     Next i
     
End Sub


'################################
'行削除
'チェックボックスがONの行を削除する
'################################
Sub 行削除()

    Dim i As Integer
    Dim shp As CheckBox
    Dim rng As Range
    Dim s As String
    
    '##############
    'メンバ変数初期化
    '##############
    m_intStartColumn = 18
    m_intEndCoumn = 0
    m_lngStartRow = 4
    m_lngEndRow = Cells(Rows.Count, 2).End(xlUp).Row
    m_lngEndRow = Range("B" & Rows.Count).End(xlUp).Row 'お試し
    m_intStartColumnRuiseki = 0
    m_intEndCoumnRuiseki = 0


    If TypeName(Selection) <> "Range" Then
        'セル以外が選択されているときに実行時エラーが発生しないように終了する
        MsgBox "セル以外が選択されている可能性があります。" & vbCrLf & "任意のセルを選択した状態にしておいて下さい。"
        Exit Sub
    Else
      For Each shp In ActiveSheet.CheckBoxes
        If shp.Value = 1 Then
            'チェックボックスONのとき
            s = shp.TopLeftCell.Address
            If s <> "$B$1" Then
                shp.Delete
                Range(s).EntireRow.Delete
            End If
        End If
      Next
    End If
    
End Sub


'################################
'行Copy
'明細の最終行からレイアウトをコピーする
'################################
Sub 行Copy()
    m_str呼出し元処理 = "行Copy"
    行数指定フォーム.Show

End Sub


'################################
'行Copy
'行数指定フォームから復帰
'################################
Sub 行コピー(p_int行数 As Integer)
    Application.ScreenUpdating = False  '画面更新の非表示
    
    Dim i As Integer
    Dim s As String
    ReDim l_arrHiduke(2) As Date
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_intNissu As Integer

    '##############
    'メンバ変数初期化
    '##############
    m_intStartColumn = 18
    m_lngStartRow = 4
    m_lngEndRow = Cells(Rows.Count, 3).End(xlUp).Row
    m_lngEndRow = Range("B" & Rows.Count).End(xlUp).Row 'お試し

    '共通チェック（現在の最終行確認含む）
    If 共通チェック = False Then
        Exit Sub
    End If
    
    '共通チェック後、改めてメンバ変数初期化
    'm_intEndCoumnRuiseki = Cells(m_lngEndRow, Columns.Count).End(xlToLeft).Column
    m_intEndCoumnRuiseki = Cells(1, Columns.Count).End(xlToLeft).Column '日次明細列の有無確認
    
    m_intStartColumnRuiseki = m_intEndCoumnRuiseki - 3
    m_intEndCoumn = m_intEndCoumnRuiseki - 4

    'チェック（日次明細列の有無）
    If m_intStartColumnRuiseki <= 17 Then
        MsgBox "日次明細列がまだ作成されていないようです。" & vbCrLf & "先に進捗管理ボタンを実行して下さい。"
        Exit Sub
    End If
    
    Range(Cells(m_lngEndRow, 1), Cells(m_lngEndRow, 1)).EntireRow.Copy
    Range(Cells(m_lngEndRow + 1, 1), Cells(m_lngEndRow + p_int行数, 1)).EntireRow.PasteSpecial
    Range(Cells(m_lngEndRow + 1, 1), Cells(m_lngEndRow + p_int行数, m_intEndCoumn)).ClearContents 'C列がブランクだと明細行(小タイトル)の背景色を設定してしまうので、"行コピー"をセット、D列以降の値をクリア
    Range(Cells(m_lngEndRow + 1, 3), Cells(m_lngEndRow + p_int行数, 3)).Value = "行コピー"
    
    m_lngEndRow = m_lngEndRow + p_int行数


    '###########################
    '日次合計/累積項目を設定する
    '###########################
    Call 日次合計_累積項目追加
    
    
    '################################
    '最小の日付、最大の日付を取得する
    '################################
    l_arrHiduke = 日付取得()
    l_datHidukeMin = l_arrHiduke(0) '最小日付
    l_datHidukeMax = l_arrHiduke(1) '最大日付
    
    '最小日付と最大日付の間の日数を取得する
    l_intNissu = DateDiff("d", l_datHidukeMin, l_datHidukeMax)


    '##############
    'フォーマット用背景色設定
    '##############
    '一時保管用変数
    Dim l_lngWkSR As Long
    Dim l_lngWkSC As Long
    Dim l_lngWkER As Long
    Dim l_lngWkEC As Long
    
    l_lngWkSR = m_lngStartRow
    l_lngWkSC = m_intStartColumn
    l_lngWkER = m_lngEndRow
    l_lngWkEC = m_intEndCoumn
    
    m_lngStartRow = 4
    m_intStartColumn = 1
    m_intEndCoumn = 17
    
    Call フォーマット用背景色設定
    
    '元のメンバ変数の値を戻す
    m_lngStartRow = l_lngWkSR
    m_intStartColumn = l_lngWkSC
    m_intEndCoumn = l_lngWkEC
    
    
    '##############
    '罫線描写
    '##############
    Call 罫線描写(l_intNissu)
    

    '##############
    '表示形式設定
    '##############
    Call 表示形式設定


    '##############
    '背景色設定
    '##############
    Call 背景色設定(l_intNissu)


    '##############
    'データの入力規則設定
    '##############
    Call データ入力規則設定
    
    
    '###########################
    'チェックボックス作成
    '###########################
    m_lngStartRow = m_lngEndRow - p_int行数 + 1   'コピー行だけチェックボックスを作成する
    
    Call CheckBox作成

    Application.ScreenUpdating = True  '画面更新の表示
End Sub


'##############
'罫線描写
'##############
Private Sub 罫線描写(p_intNissu As Integer)

    Dim l_lngWrkCellY As Long
    Dim j As Integer
    Dim l_intNissu As Integer
    
    '変数初期化
    l_lngWrkCellY = m_intStartColumn
    l_intNissu = p_intNissu
    
    '罫線描写
    Range(Cells(2, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders.LineStyle = xlContinuous
    
    '太線描写
    For j = m_intStartColumn To l_intNissu + m_intStartColumn + 1   '"+1"は累積項目分
        Range(Cells(1, l_lngWrkCellY), Cells(m_lngEndRow + 2, l_lngWrkCellY + 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Range(Cells(1, l_lngWrkCellY), Cells(m_lngEndRow + 2, l_lngWrkCellY + 3)).Borders(xlEdgeRight).Weight = xlMedium
        Range(Cells(1, l_lngWrkCellY), Cells(m_lngEndRow + 2, l_lngWrkCellY + 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Range(Cells(1, l_lngWrkCellY), Cells(m_lngEndRow + 2, l_lngWrkCellY + 3)).Borders(xlEdgeRight).Weight = xlMedium
        
        l_lngWrkCellY = l_lngWrkCellY + 4
    Next j
    
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeLeft).Weight = xlMedium
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeTop).Weight = xlMedium
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeRight).Weight = xlMedium
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Cells(m_lngEndRow + 1, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders(xlEdgeBottom).Weight = xlMedium

End Sub


'##############
'表示形式設定
'##############
Private Sub 表示形式設定()

    '日付形式を整える
    Range(Cells(1, m_intStartColumn), Cells(1, m_intEndCoumn)).NumberFormatLocal = "mm月dd日"
    
    '太字にする
    Range(Cells(m_lngStartRow, m_intStartColumnRuiseki), Cells(m_lngEndRow + 1, m_intEndCoumnRuiseki)).Font.Bold = True
    
End Sub


'##############
'日次明細列及びタイトル行の背景色設定
'##############
Private Sub 背景色設定(p_intNissu As Integer)

    Dim j As Integer
    Dim k As Integer
    Dim i As Integer
    Dim l_intNissu As Integer
    Dim l_lngWrkCellY As Long
    Dim l_lngWrkCellYWeek As Long

    '変数初期化
    l_lngWrkCellY = m_intStartColumn
    l_lngWrkCellYWeek = m_intStartColumn
    l_intNissu = p_intNissu
    
    For j = m_intStartColumn To l_intNissu + m_intStartColumn + 1     '列単位の処理。 ""+1"は累積項目分
        '明細行の背景色設定
        For k = 0 To 3                  '1日4列分の処理
            'If ((l_lngWrkCellY + k) Mod 2) = 1 Then
            If ((l_lngWrkCellY + k) Mod 4) = 2 Then                   '実施完了-予定
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 40
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 3 Then               '実施完了-実績
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 24
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 0 Then               '精査完了-予定
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 35
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 1 Then               '精査完了-実績
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 20
            Else
                '処理なし
            End If
        Next k
        
        'タイトル行の背景色設定
        If l_lngWrkCellYWeek < m_intEndCoumn Then
            If (j Mod 2) = 0 Then '1週間分（28列分)毎の処理
                Range(Cells(1, l_lngWrkCellYWeek), Cells(3, l_lngWrkCellYWeek + 27)).Interior.ColorIndex = 42
            Else
                Range(Cells(1, l_lngWrkCellYWeek), Cells(3, l_lngWrkCellYWeek + 27)).Interior.ColorIndex = 34
            End If
        End If
        l_lngWrkCellY = l_lngWrkCellY + 4
        l_lngWrkCellYWeek = l_lngWrkCellYWeek + 28
    Next j
    
    '明細行(小タイトル)の背景色設定
    For i = 4 To m_lngEndRow       '行単位の処理
            If Cells(i, 2).Value = "" And Cells(i, 5).Value = "" And Cells(i, 7).Value = "" And Cells(i, 9).Value = "" And Cells(i, 12).Value = "" Then
                Range(Cells(i, 1), Cells(i, m_intEndCoumn + 4)).Interior.ColorIndex = 33
            End If
    Next i
    
    '累積項目タイトル行の背景色設定
    Range(Cells(1, m_intStartColumnRuiseki), Cells(3, m_intEndCoumnRuiseki)).Interior.ColorIndex = 44
    
    '累積項目の合計行の背景色設定
    Range(Cells(m_lngEndRow + 1, m_intStartColumnRuiseki), Cells(m_lngEndRow + 1, m_intEndCoumnRuiseki)).Interior.ColorIndex = 27

End Sub


'##############
'タイトル設定
'##############
Private Function タイトル設定(p_intNissu As Integer, p_datHidukeMin As Date) As Long

    Dim j As Integer
    Dim k As Integer
    Dim i As Integer
    Dim l_lngWrkCellY As Long
    Dim l_intNissu As Integer
    Dim l_datHidukeMin As Date

    
    '変数初期化
    l_lngWrkCellY = m_intStartColumn
    l_intNissu = p_intNissu
    l_datHidukeMin = p_datHidukeMin
    
    For j = m_intStartColumn To l_intNissu + m_intStartColumn   '列単位の処理
        For k = 0 To 3                                          '1日4列分の処理
            If ((l_lngWrkCellY + k) Mod 4) = 2 Then             '実施完了-予定
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "実施完了"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "予定"
                    End If
                Next i
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 3 Then         '実施完了-実績
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "実施完了"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "実績"
                    End If
                Next i
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 0 Then         '精査完了-予定
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "精査完了"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "予定"
                    End If
                Next i
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 1 Then         '精査完了-実績
                For i = 1 To m_lngEndRow                        '行単位の処理
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "精査完了"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "実績"
                    End If
                Next i
            Else
                '処理なし
            End If
        Next k
        
        l_lngWrkCellY = l_lngWrkCellY + 4
    Next j
    
    タイトル設定 = l_lngWrkCellY
    
End Function


Sub 初期化()

    m_str呼出し元処理 = "初期化"
    行数指定フォーム.Show

End Sub
 
Public Sub フォーマット化(p_int行数 As Integer)
    Application.ScreenUpdating = False  '画面更新の非表示
    
    Dim i As Integer
    Dim s As String
    ReDim l_arrHiduke(2) As Date
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_intNissu As Integer
    

    '##############
    'メンバ変数初期化
    '##############
    m_intStartColumn = 1
    m_intEndCoumn = 17
    m_lngStartRow = 4
    m_lngEndRow = 4 + p_int行数 - 1


    '################################
    'フォーマット用クリア処理
    '################################
    フォーマット用クリア処理


    '##############
    '罫線描写
    '##############
    Call フォーマット用罫線描写
    

    '##############
    '表示形式設定
    '##############
    Call フォーマット用表示形式設定


    '##############
    'フォーマット用背景色設定
    '##############
    Call フォーマット用背景色設定
    
    
    '##############
    'データの入力規則設定
    '##############
    Call データ入力規則設定

    
    '###########################
    'チェックボックス作成
    '###########################
    Call CheckBox作成
    
    MsgBox "初期化完了"

    Application.ScreenUpdating = True  '画面更新の表示
End Sub

'#######################
'フォーマット用罫線描写
'#######################
Private Sub フォーマット用罫線描写()
    
    '罫線描写
    Range(Cells(m_lngStartRow, m_intStartColumn), Cells(m_lngEndRow, m_intEndCoumn)).Borders.LineStyle = xlContinuous
    
End Sub

'###########################
'フォーマット用表示形式設定
'###########################
Private Sub フォーマット用表示形式設定()

    '日付形式を整える
    Range(Cells(m_lngStartRow, 7), Cells(m_lngEndRow, 10)).NumberFormatLocal = "mm月dd日"
    Range(Cells(m_lngStartRow, 12), Cells(m_lngEndRow, 13)).NumberFormatLocal = "mm月dd日"
    Range(Cells(m_lngStartRow, 15), Cells(m_lngEndRow, 16)).NumberFormatLocal = "mm月dd日"

    '配置を整える
    Cells.HorizontalAlignment = xlCenter
    Range(Cells(m_lngStartRow, 3), Cells(m_lngEndRow, 3)).HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter
    
End Sub


'##########################
'フォーマット用背景色設定
'##########################
Private Sub フォーマット用背景色設定()
    
    Range(Cells(m_lngStartRow, m_intStartColumn), Cells(m_lngEndRow, m_intEndCoumn)).Interior.ColorIndex = 2
    
    '入力セルについて背景色を設定する
    Range(Cells(m_lngStartRow, 2), Cells(m_lngEndRow, 2)).Interior.ColorIndex = 36       'ID
    Range(Cells(m_lngStartRow, 3), Cells(m_lngEndRow, 3)).Interior.ColorIndex = 19      'シナリオ名
    '19/04/04 Add Start
    Range(Cells(m_lngStartRow, 4), Cells(m_lngEndRow, 3)).Interior.ColorIndex = 19      '機能名
    '19/04/04 Add End
    Range(Cells(m_lngStartRow, 5), Cells(m_lngEndRow, 5)).Interior.ColorIndex = 19      'ケース数
    Range(Cells(m_lngStartRow, 6), Cells(m_lngEndRow, 6)).Interior.ColorIndex = 19      '実施者
    Range(Cells(m_lngStartRow, 7), Cells(m_lngEndRow, 7)).Interior.ColorIndex = 36       'テスト実施-着手-予定
    Range(Cells(m_lngStartRow, 9), Cells(m_lngEndRow, 9)).Interior.ColorIndex = 36       'テスト実施-完了-予定
    Range(Cells(m_lngStartRow, 11), Cells(m_lngEndRow, 11)).Interior.ColorIndex = 19    '精査者
    Range(Cells(m_lngStartRow, 12), Cells(m_lngEndRow, 12)).Interior.ColorIndex = 36     'テスト精査-完了-予定
    Range(Cells(m_lngStartRow, 14), Cells(m_lngEndRow, 14)).Interior.ColorIndex = 19    '指摘有無
    Range(Cells(m_lngStartRow, 15), Cells(m_lngEndRow, 15)).Interior.ColorIndex = 19    '指摘対応-(実施者)指摘修正日
    Range(Cells(m_lngStartRow, 16), Cells(m_lngEndRow, 16)).Interior.ColorIndex = 19    '指摘対応-(精査者)指摘確認日
    Range(Cells(m_lngStartRow, 17), Cells(m_lngEndRow, 17)).Interior.ColorIndex = 19    '状況
End Sub


'##########################
'データ入力規則設定
'##########################
Private Sub データ入力規則設定()

    Range(Cells(m_lngStartRow, 14), Cells(m_lngEndRow, 14)).Select
    With Range(Cells(m_lngStartRow, 14), Cells(m_lngEndRow, 14)).Validation
            .Delete
            .Add Type:=xlValidateList, _
                Formula1:="有,無"
    End With
    
    Range(Cells(m_lngStartRow, 17), Cells(m_lngEndRow, 17)).Select
    With Range(Cells(m_lngStartRow, 17), Cells(m_lngEndRow, 17)).Validation
            .Delete
            .Add Type:=xlValidateList, _
                Formula1:="実施中,精査中,再実施中,再精査中,完了"
    End With

End Sub


'################################
'フォーマット用クリア処理
'################################
Public Sub フォーマット用クリア処理()

    Dim shp As CheckBox

    Range("R1", Cells(1, Columns.Count)).EntireColumn.Delete
    Range("A4", Cells(Rows.Count, 2)).EntireRow.Delete
    
    'チェックボックスを削除する
    'ActiveSheet.CheckBoxes.Delete
    For Each shp In ActiveSheet.CheckBoxes
        If shp.TopLeftCell.Address <> "$B$1" Then
                shp.Delete                          '日次明細列の値削除を確認するチェックボックス以外は削除する
        End If
    Next

End Sub


'################################
'共通チェック
'################################
Public Function 共通チェック() As Boolean

    '変数宣言
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_datwkHiduke  As Date
    Dim i As Integer
    
    '変数初期化
    共通チェック = True
    l_datHidukeMin = "9999/12/31"
    l_datHidukeMax = "1900/01/01"
    
    
    '<チェック>
    '①コピー対象行の存在チェック
    If m_lngEndRow <= 3 Then
        MsgBox "4行目以降にIDを入力して下さい。"
        
        共通チェック = False
        Exit Function
    End If
    
    
    '② 予定日の入力チェック
    '(テスト実施) 最小の着手予定日を取得する
    For i = m_lngStartRow To m_lngEndRow
        If Cells(i, 7) = "" Then
        Else
            l_datwkHiduke = Cells(i, 7).Value
            If l_datwkHiduke < l_datHidukeMin Then
                l_datHidukeMin = l_datwkHiduke
            End If
        End If
    Next i
    
    If l_datHidukeMin = "9999/12/31" Then
        MsgBox "テスト実施-着手予定日を最低１つは入力して下さい。"
        
        共通チェック = False
        Exit Function
    End If
    
    '(テスト実施) 最大の完了予定日を取得する
    For i = m_lngStartRow To m_lngEndRow
        If Cells(i, 9) = "" Then
        Else
            l_datwkHiduke = Cells(i, 9).Value
        If l_datwkHiduke > l_datHidukeMax Then
            l_datHidukeMax = l_datwkHiduke
            End If
        End If
    Next i

    If l_datHidukeMax = "1900/01/01" Then
        MsgBox "テスト実施-完了予定日を最低１つは入力して下さい。"
        
        共通チェック = False
        Exit Function
    End If
    
    '(テスト精査) 最大の完了予定日を取得する
    l_datHidukeMax = "1900/01/01"
    
    For i = m_lngStartRow To m_lngEndRow
        If Cells(i, 12) = "" Then
        Else
            l_datwkHiduke = Cells(i, 12).Value
            If l_datwkHiduke > l_datHidukeMax Then
                l_datHidukeMax = l_datwkHiduke
            End If
        End If
    Next i

    If l_datHidukeMax = "1900/01/01" Then
        MsgBox "テスト精査-完了予定日を最低１つは入力して下さい。"
        
        共通チェック = False
        Exit Function
    End If
    
End Function


'バックアップしたシートから値を再設定する
'################################
Private Sub 値再設定(p_datHidukeMin As Date, p_datHidukeMax As Date)


    Worksheets("bk_日次管理").Activate
    
    Dim l_intLastColumn_bk日次管理 As Integer
    Dim l_datHidukeMin_bk日次管理 As Date
    Dim l_datHidukeMax_bk日次管理 As Date
    Dim l_intNissu_bk日次管理 As Integer
        
    'bk_日次管理シート上のセルR1の最小日付を取得
    l_datHidukeMin_bk日次管理 = Cells(1, 18).Value
    
    'bk_日次管理シート上の最大日付を取得
    l_intLastColumn_bk日次管理 = Cells(1, Columns.Count).End(xlToLeft).Column       '最終列
    l_datHidukeMax_bk日次管理 = Cells(1, l_intLastColumn_bk日次管理 - 4)            '最終列から累積4項目を除く
    
    '日次管理シートの最小日付とbk_日次管理シート上の最小日付を比較する
    If p_datHidukeMin < l_datHidukeMin_bk日次管理 Then
        '日次管理シートの最小日付の方が小さいとき、bk_日次管理シートの最小日付から値を取得しても、日次管理シート上で貼り付け先セルは存在する
        'よって、bk_日次管理シート上の最小日付を設定する
        l_datHidukeMin_bk日次管理 = l_datHidukeMin_bk日次管理                       '特に処理に意味なし
    Else
        'bk_日次管理シートの最小日付の方が小さいとき、bk_日次管理シートの最小日付から値を取得すると、日次管理シート上で貼り付け先セルは存在しないことになる
        'よって、日次管理シート上の最小日付を設定する
        l_datHidukeMin_bk日次管理 = p_datHidukeMin
    End If
    
    '日次管理シートの最大日付とbk_日次管理シート上の最大日付を比較する
    If p_datHidukeMax > l_datHidukeMax_bk日次管理 Then
        '日次管理シートの最大日付の方が大きいとき、bk_日次管理シートの最大日付から値を取得しても、日次管理シート上で貼り付け先セルは存在する
        'よって、bk_日次管理シート上の最大日付を設定する
        If l_datHidukeMax_bk日次管理 = "0:00:00" Then                               '但し、初回の進捗管理ボタン押下時はbk_日次管理シート上に日次明細列がないので、日付は日次管理シート上から取得する
            l_datHidukeMax_bk日次管理 = p_datHidukeMax
        Else
            l_datHidukeMax_bk日次管理 = l_datHidukeMax_bk日次管理                   '特に処理に意味なし
        End If
    Else
        'bk_日次管理シートの最大日付の方が大きいとき、bk_日次管理シートの最大日付から値を取得すると、日次管理シート上で貼り付け先セルは存在しないことになる
        'よって、日次管理シート上の最大日付を設定する
        l_datHidukeMax_bk日次管理 = p_datHidukeMax
    End If
    
    '最小日付と最大日付の間の日数を取得する
    l_intNissu_bk日次管理 = DateDiff("d", l_datHidukeMin_bk日次管理, l_datHidukeMax_bk日次管理)     '(例)20日～26日のカウントは6日間。（つまり開始日(20日)はカウントしていない）
    l_intNissu_bk日次管理 = l_intNissu_bk日次管理 + 1                                               'そのため、＋1日する


    '############################
    'bk_日次管理シート上の管理明細の値を日次管理シート上の同じ日付の項目にセットする
    '############################
    '日次管理シート上で上記で見つけた最小日付をセットしている列項目を探索する
    Worksheets("日次管理").Activate
    
    Dim j As Integer
    Dim l_intLastColumn_日次管理 As Integer
    Dim l_datHidukeMin_日次管理 As Date
    Dim l_datHidukeMax_日次管理 As Date
    Dim l_intNissu_日次管理 As Integer
    
    l_datHidukeMin_日次管理 = Cells(1, 18).Value
    l_intLastColumn_日次管理 = Cells(1, Columns.Count).End(xlToLeft).Column '最終列
    l_datHidukeMax_日次管理 = Cells(1, l_intLastColumn_日次管理 - 4)                        '最終列から累積4項目を除く
    
    '最小日付と最大日付の間の日数を取得する
    l_intNissu_日次管理 = DateDiff("d", l_datHidukeMin_日次管理, l_datHidukeMax_日次管理)   '(例)20日～26日のカウントは6日間。（つまり開始日(20日)はカウントしていない）
    l_intNissu_日次管理 = l_intNissu_日次管理 + 1                                           'そのため、＋1日する
    
    j = 0
    
    For j = m_intStartColumn To l_intNissu_日次管理 * 4 + m_intStartColumn + 1              '列単位の処理。 ""+1"は累積項目分
        'bk_日次管理シート上の最小日付と日次管理シート上の日付が一致するセルを確認
        If l_datHidukeMin_bk日次管理 = Range(Cells(1, j), Cells(1, j)).Value Then
            '日次管理シート上の最小日付が見つかった場合、bk_日次管理シート上から値をセル範囲一括でコピー＆ペースト
            Worksheets("bk_日次管理").Activate
            Worksheets("bk_日次管理").Range(Cells(m_lngStartRow, m_intStartColumn), Cells(m_lngEndRow, m_intStartColumn - 1 + (l_intNissu_bk日次管理 * 4))).Copy
                       
            Worksheets("日次管理").Activate
            ActiveSheet.Range(ActiveSheet.Cells(m_lngStartRow, j), ActiveSheet.Cells(m_lngEndRow, j - 1 + (l_intNissu_bk日次管理 * 4))).PasteSpecial Paste:=xlPasteValues
            
            Exit For
        End If
    Next j

End Sub


'################################
'予実管理シートを作成する
'################################
Private Sub 予実管理作成(p_intNissu As Integer)

    Worksheets("予実管理").Activate

    'シート初期化
    Application.DisplayAlerts = False   '確認メッセージの非表示
    Range("3:3").UnMerge                'セルの結合解除
    Cells.ClearFormats                  '全セルの表示形式クリア
    Cells.Clear                         '全セルのクリア
    Cells.ColumnWidth = 5               '全セルの幅
    Cells.RowHeight = 20                '全セルの高さ
    '19/03/26 Add Start
    ActiveSheet.Columns.ClearOutline    '列のグループ化を全て削除
    '19/03/26 Add End

    '変数宣言
    Dim Dic担当者 As Object
    Dim Dic精査者 As Object
    Dim l_intテスト実施start As Integer
    Dim l_intテスト精査start As Integer
    Dim i As Integer
    
    'タイトル
    With Cells(1, 1)
        .Value = "予実管理"
        .Font.Bold = True
    End With
    '19/03/26 DEL Start
'    With Cells(1, 2)
'        .Value = "(金曜日を週末基準とする)"
'        .Font.Bold = True
'    End With
    '19/03/26 DEL End

    '###########################
    '## テスト実施 ##
    '###########################
    'テスト実施の開始行の設定
    l_intテスト実施start = 2
    
    'テスト実施者の取得
    Set Dic担当者 = メンバー取得("担当者")
    
    'テスト実施者の予実管理表作成
    Call 予実管理明細表作成(l_intテスト実施start, Dic担当者, p_intNissu)


    '###########################
    '## テスト精査 ##
    '###########################
    'テスト精査の開始行の設定
    l_intテスト精査start = 5 + Dic担当者.Count + 3

    'テスト精査者の取得
    Set Dic精査者 = メンバー取得("精査者")

    'テスト精査者の予実管理表作成
    Call 予実管理明細表作成(l_intテスト精査start, Dic精査者, p_intNissu)

    Cells.HorizontalAlignment = xlCenter
    Cells(1, 2).HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter

        
    Application.DisplayAlerts = True    '確認メッセージの表示
    
    Worksheets("日次管理").Activate

End Sub


Private Function メンバー取得(p_strRoll As String) As Object

    Dim Dic As Object
    Dim i As Integer
    Dim l_strName As String
    Set Dic = CreateObject("Scripting.Dictionary")
    
    Worksheets("日次管理").Activate
    
    For i = m_lngStartRow To m_lngEndRow
        If p_strRoll = "担当者" Then
            '担当者を取得
            l_strName = Cells(i, 6).Value
        Else
            '精査者を取得
            l_strName = Cells(i, 11).Value
        End If
        If Not l_strName = "" Then
            If Not Dic.exists(l_strName) Then
                Dic.Add l_strName, Null
            End If
        End If
    Next i

    Set メンバー取得 = Dic
    
    Worksheets("予実管理").Activate
    
End Function


'###########################
'実績日を登録する
'###########################
Private Sub 実績日登録()

    '変数宣言
    Dim i As Integer
    Dim j As Integer
    
    For i = m_lngStartRow To m_lngEndRow
    
        '19/03/26 Add Start
        'テスト実施-着手-実績値の登録
        For j = m_intStartColumn To m_intEndCoumn Step 1        '日次明細列を昇順に繰り回し
            If (j Mod 4) = 3 Then                               '実施完了-実績の列
                If Cells(i, j).Value <> "" Then
                    Cells(i, 8).Value = Cells(1, j).Value
                    Exit For
                End If
            End If
        Next j
        '19/03/26 Add End
    
        '累積実施完了-予定数/実績数の一致を確認
        If Cells(i, m_intStartColumnRuiseki).Value = Cells(i, m_intStartColumnRuiseki + 1).Value Then
            'テスト実施-完了-実績値の登録
            For j = m_intEndCoumn To m_intStartColumn Step -1       '日次明細列を降順に繰り回し
                If (j Mod 4) = 3 Then                               '実施完了-実績の列
                    If Cells(i, j).Value <> "" Then
                        Cells(i, 10).Value = Cells(1, j).Value
                        Exit For
                    End If
                End If
            Next j

            'テスト実施-着手-実績値の登録
            For j = m_intStartColumn To m_intEndCoumn Step 1        '日次明細列を昇順に繰り回し
                If (j Mod 4) = 3 Then                               '実施完了-実績の列
                    If Cells(i, j).Value <> "" Then
                        Cells(i, 8).Value = Cells(1, j).Value
                        Exit For
                    End If
                End If
            Next j
                
        End If
        
        '累積精査完了-予定数/実績数の一致を確認
        If Cells(i, m_intStartColumnRuiseki + 2).Value = Cells(i, m_intStartColumnRuiseki + 3).Value Then
            'テスト精査-完了-実績値の登録
            For j = m_intEndCoumn To m_intStartColumn Step -1       '日次明細列を降順に繰り回し
                If (j Mod 4) = 1 Then                               '精査完了-実績の列
                    If Cells(i, j).Value <> "" Then
                        Cells(i, 13).Value = Cells(1, j).Value
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    

End Sub


'###########################
'予実管理明細表作成
'###########################
Private Sub 予実管理明細表作成(p_int開始行 As Integer, p_dicメンバー As Object, p_intNissu As Integer)

    '変数宣言
    Dim l_sht_日次管理 As Worksheet
    Dim l_sht_Yojitsu As Worksheet
    Dim Keys() As Variant
    Dim l_intNissu As Integer
    Dim l_wrkDate As Date
    Dim l_intWeeks As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim n As Integer
    Dim l_bolChgFrg As Boolean

    '変数初期化
    Set l_sht_日次管理 = Sheets("日次管理")
    Set l_sht_Yojitsu = Sheets("予実管理")
    
    If p_int開始行 = 2 Then     'テスト実施
        With Cells(p_int開始行, 1)
            .Value = "テスト実施"
            .Font.Bold = True
        End With
        With Cells(p_int開始行, 2)
            .Value = "(累積管理)"
            .Font.Bold = True
        End With
    Else                        'テスト精査
        With Cells(p_int開始行, 1)
            .Value = "テスト精査"
            .Font.Bold = True
        End With
        With Cells(p_int開始行, 2)
            .Value = "(累積管理)"
            .Font.Bold = True
        End With
    End If

    '本日日付
    With Cells(p_int開始行 + 1, 1)
        .Value = "=TODAY()"
        .Font.ColorIndex = 2
        .Font.Bold = True
    End With


    '##############
    'テスト実施者の取得
    '##############
    Cells(p_int開始行 + 2, 1).Value = "担当"
    
    '担当者/精査者をセット
    Keys = p_dicメンバー.Keys
   
    For j = 0 To p_dicメンバー.Count - 1
        Cells(p_int開始行 + 3 + j, 1).Value = Keys(j)
    Next j
    
    '合計
    Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1).Value = "合計"
  
     '19/03/26 Mod Start
'    '１週間単位に「予定」「実績」列項目を用意する
'    l_intNissu = p_intNissu
'    l_wrkDate = l_sht_日次管理.Cells(1, m_intStartColumn).Value
'    l_intWeeks = (m_intEndCoumn - m_intStartColumn) / 4 / 7
'
'    For k = 0 To l_intWeeks - 1
'        For m = 0 To 1
'            Cells(p_int開始行 + 1, 2 + (2 * k) + m).Value = l_wrkDate + 4   '金曜日を週末基準とする
'
'            If ((2 + m) Mod 2) = 0 Then      '予定
'                Cells(p_int開始行 + 2, 2 + (2 * k) + m).Value = "予定"
'                '担当者
'                If p_int開始行 = 2 Then     'テスト実施
'                    For j = 0 To p_dicメンバー.Count - 1
'                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 6).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """実施完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """予定""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
'                    Next j
'                Else                        'テスト精査
'                    For j = 0 To p_dicメンバー.Count - 1
'                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 11).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """精査完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """予定""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
'                    Next j
'
'                End If
'                '合計
'                Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + (2 * k) + m) = _
'                "=SUM(" + l_sht_Yojitsu.Cells(p_int開始行 + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + p_dicメンバー.Count - 1, 2 + (2 * k) + m).Address + ")"
'
'            ElseIf ((2 + m) Mod 2) = 1 Then  '実績
'                Cells(p_int開始行 + 2, 2 + (2 * k) + m).Value = "実績"
'                '担当者
'                If p_int開始行 = 2 Then     'テスト実施
'                    For j = 0 To p_dicメンバー.Count - 1
'                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 6).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """実施完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """実績""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
'                    Next j
'                Else                        'テスト精査
'                    For j = 0 To p_dicメンバー.Count - 1
'                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 11).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """精査完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """実績""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
'                    Next j
'                End If
'                '合計
'                Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + (2 * k) + m) = _
'                "=SUM(" + l_sht_Yojitsu.Cells(p_int開始行 + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + p_dicメンバー.Count - 1, 2 + (2 * k) + m).Address + ")"
'            End If
'        Next m
'        l_wrkDate = l_wrkDate + 7
'    Next k
    
'    '「実績日による管理」列項目を用意する
'    With Range(Cells(p_int開始行 + 1, 1 + l_intWeeks * 2 + 1), Cells(p_int開始行 + 1, 1 + l_intWeeks * 2 + 2))
'        .Merge
'        .Value = "実績日による管理" + vbCrLf + "(指摘有無に関わらず" + vbCrLf + "初回の実績入力日で管理)"
'        .WrapText = True
'        .ColumnWidth = 15
'        .RowHeight = 40
'    End With
'
'    Cells(p_int開始行 + 2, 1 + l_intWeeks * 2 + 1) = "進捗率"
'    '担当者 + 合計
'    For j = 0 To p_dicメンバー.Count
'        With Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2 + 1)
'            .Value = "=SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1 + l_intWeeks * 2).Address + ")*(" + """実績""" + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 1 + l_intWeeks * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2).Address + ")" + _
'            "/SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1 + l_intWeeks * 2).Address + ")*(" + """予定""" + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 1 + l_intWeeks * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2).Address + ")"
'            .NumberFormatLocal = "0%"
'        End With
'    Next j
'
'    Cells(p_int開始行 + 2, 1 + l_intWeeks * 2 + 2) = "完了率"
'    '担当者 + 合計
'    For j = 0 To p_dicメンバー.Count
'        With Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2 + 2)
'            .Value = "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2).Address + "/" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2 - 1).Address
'            .NumberFormatLocal = "0%"
'        End With
'    Next j
'
'    '背景色
'    Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 1, 1 + l_intWeeks * 2 + 2)).Interior.ColorIndex = 40
'    Range(Cells(p_int開始行 + 2, 1), Cells(p_int開始行 + 2, 1 + l_intWeeks * 2 + 2)).Interior.ColorIndex = 38
'    '担当者 + 合計
'    For j = 0 To p_dicメンバー.Count
'        If (j Mod 2) = 1 Then
'            Range(Cells(p_int開始行 + 3 + j, 1), Cells(p_int開始行 + 3 + j, 1 + l_intWeeks * 2 + 2)).Interior.ColorIndex = 24
'        End If
'    Next j
'
'    '罫線描写
'    Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1 + l_intWeeks * 2 + 2)).Borders.LineStyle = xlContinuous
'
'    '太線描写
'    '外枠
'    With Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1 + l_intWeeks * 2 + 2))
'        .Borders(xlEdgeRight).Weight = xlMedium
'        .Borders(xlEdgeLeft).Weight = xlMedium
'        .Borders(xlEdgeTop).Weight = xlMedium
'        .Borders(xlEdgeBottom).Weight = xlMedium
'    End With
'    '内訳
'    Range(Cells(p_int開始行 + 2, 1), Cells(p_int開始行 + 2, 1 + l_intWeeks * 2 + 2)).Borders(xlEdgeBottom).Weight = xlMedium
'
'    For k = 0 To l_intWeeks * 2
'        If (k Mod 2) = 0 Then
'            With Range(Cells(p_int開始行 + 1, 2 + k), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + k))
'                .Borders(xlEdgeLeft).Weight = xlMedium
'            End With
'        End If
'    Next k
'
'
'    Rows(3).Columns.AutoFit                                         'セルの列幅自動設定
'
'    For n = 1 To l_intWeeks * 2 + 1
'        Cells(3, n).ColumnWidth = Cells(3, n).ColumnWidth * 1.2     'セルの列幅自動設定 ×１．２倍
'    Next n


    '１日単位に「予定」「実績」列項目を用意する
    l_intNissu = p_intNissu + 1
    l_wrkDate = l_sht_日次管理.Cells(1, m_intStartColumn).Value
    l_intWeeks = (m_intEndCoumn - m_intStartColumn) / 4 / 7

    For k = 0 To l_intNissu - 1
        For m = 0 To 1
            Cells(p_int開始行 + 1, 2 + (2 * k) + m).Value = l_wrkDate   '基準日

            If ((2 + m) Mod 2) = 0 Then     '予定
                Cells(p_int開始行 + 2, 2 + (2 * k) + m).Value = "予定"
                '担当者
                If p_int開始行 = 2 Then     'テスト実施
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 6).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """実施完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """予定""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
                    Next j
                Else                        'テスト精査
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 11).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """精査完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """予定""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
                    Next j

                End If
                '合計
                Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int開始行 + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + p_dicメンバー.Count - 1, 2 + (2 * k) + m).Address + ")"
                
            ElseIf ((2 + m) Mod 2) = 1 Then  '実績
                Cells(p_int開始行 + 2, 2 + (2 * k) + m).Value = "実績"
                '担当者
                If p_int開始行 = 2 Then     'テスト実施
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 6).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """実施完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """実績""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
                    Next j
                Else                        'テスト精査
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 11).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """精査完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """実績""" + "),日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！
                    Next j
                End If
                '合計
                Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int開始行 + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + p_dicメンバー.Count - 1, 2 + (2 * k) + m).Address + ")"
            End If
        Next m
        l_wrkDate = l_wrkDate + 1
    Next k

    '「実績日による管理」列項目を用意する
    With Range(Cells(p_int開始行 + 1, 1 + l_intNissu * 2 + 1), Cells(p_int開始行 + 1, 1 + l_intNissu * 2 + 2))
        .Merge
        .Value = "実績日による管理" + vbCrLf + "(指摘有無に関わらず" + vbCrLf + "初回の実績入力日で管理)"
        .WrapText = True
        .ColumnWidth = 15
        .RowHeight = 40
    End With
    
    Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 1) = "進捗率"
    '担当者 + 合計
    For j = 0 To p_dicメンバー.Count
        With Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 + 1)
            .Value = "=SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1 + l_intNissu * 2).Address + ")*(" + """実績""" + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2).Address + ")" + _
            "/SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1 + l_intNissu * 2).Address + ")*(" + """予定""" + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2).Address + ")"
            .NumberFormatLocal = "0%"
        End With
    Next j

    Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2) = "完了率"
    '担当者 + 合計
    For j = 0 To p_dicメンバー.Count
        With Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 + 2)
            .Value = "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2).Address + "/" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 - 1).Address
            .NumberFormatLocal = "0%"
        End With
    Next j
    
    '背景色
    'Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 40
    'Range(Cells(p_int開始行 + 2, 1), Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 38
    
    For k = 1 To l_intNissu * 2 + 1 Step 2
        With Range(Cells(p_int開始行 + 1, 1 + k), Cells(p_int開始行 + 2, 2 + k))
        
        If l_bolChgFrg Then
            .Interior.ColorIndex = 35
            l_bolChgFrg = False
        Else
            .Interior.ColorIndex = 40
            l_bolChgFrg = True
        End If
        
        End With
    Next k
    
    Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 2, 1)).Interior.ColorIndex = 38
    
        
    '担当者 + 合計
    For j = 0 To p_dicメンバー.Count
        If (j Mod 2) = 1 Then
            Range(Cells(p_int開始行 + 3 + j, 1), Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 24
        End If
    Next j

    '罫線描写
    Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1 + l_intNissu * 2 + 2)).Borders.LineStyle = xlContinuous
    
    '太線描写
    '外枠
    With Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1 + l_intNissu * 2 + 2))
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    '内訳
    Range(Cells(p_int開始行 + 2, 1), Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2)).Borders(xlEdgeBottom).Weight = xlMedium
    
    For k = 0 To l_intNissu * 2
        If (k Mod 2) = 0 Then
            With Range(Cells(p_int開始行 + 1, 2 + k), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + k))
                .Borders(xlEdgeLeft).Weight = xlMedium
            End With
        End If
    Next k


    Rows(3).Columns.AutoFit                                         'セルの列幅自動設定
    
    For n = 1 To l_intNissu * 2 + 1
        Cells(3, n).ColumnWidth = Cells(3, n).ColumnWidth * 1.2     'セルの列幅自動設定 ×１．２倍
    Next n

    '列のグループ化
    If p_int開始行 <> 2 Then    '精査の表を作る時のみ実施
        For k = 0 To l_intNissu * 2
            If (k Mod 14) = 0 And k <> 0 Then
                With Range(Cells(p_int開始行 + 1, k - 13 + 1), Cells(p_int開始行 + 1, k - 1))
                    .Columns.Group                                      '非表示状態にしたい個所をグループ化
                    ActiveSheet.Outline.ShowLevels columnlevels:=1      '現時点でグループ化されている個所を非表示にする
                End With
            End If
        Next k
    End If
    
    
    '19/03/26 Mod End
    
    
End Sub

'19/04/04 Add Start
'################################
'予実管理(機能単位)シートを作成する
'################################
Private Sub 予実管理作成_機能単位(p_intNissu As Integer)

    Worksheets("予実管理_機能単位").Activate

    'シート初期化
    Application.DisplayAlerts = False   '確認メッセージの非表示
    Range("3:3").UnMerge                'セルの結合解除
    Cells.ClearFormats                  '全セルの表示形式クリア
    Cells.Clear                         '全セルのクリア
    Cells.ColumnWidth = 5               '全セルの幅
    Cells.RowHeight = 20                '全セルの高さ
    ActiveSheet.Columns.ClearOutline    '列のグループ化を全て削除

    '変数宣言
    Dim Dic担当者 As Object
    Dim Dic精査者 As Object
    Dim Dic機能 As Object
    Dim Dic対象機能 As Object
    Dim l_intテスト実施start As Integer
    Dim l_intテスト精査start As Integer
    Dim l_int機能メンバー数 As Integer
    Dim l_int機能_担当者数 As Integer
    Dim l_int機能_精査者数 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Keys_機能() As Variant
    Dim Keys_担当者() As Variant
    Dim Keys_精査者() As Variant
    Dim l_str機能 As String
    
    Set Dic対象機能 = CreateObject("Scripting.Dictionary")
    
    'タイトル
    With Cells(1, 1)
        .Value = "予実管理(機能単位)"
        .Font.Bold = True
    End With


    '機能の取得
    Set Dic機能 = 機能取得()
    Keys_機能 = Dic機能.Keys                    '機能キーをセット
    
    'テスト実施者の取得
    Set Dic担当者 = メンバー取得_機能単位("担当者")
    Keys_担当者 = Dic担当者.Keys                '機能キーをセット
    
    'テスト精査者の取得
    Set Dic精査者 = メンバー取得_機能単位("精査者")
    Keys_精査者 = Dic精査者.Keys                '機能キーをセット

    
    '機能単位にテスト実施、テスト精査の予実管理表を作成する。
    For j = 0 To Dic機能.Count - 1
    
        '機能名の取得
        l_str機能 = Keys_機能(j)
            
        '最初の機能によるテスト実施の開始行の設定
        If j = 0 Then
            l_intテスト実施start = 2
        End If
                
                
        '################
        '## テスト実施 ##
        '################

        '対象機能の担当者を抽出する。（次のテスト精査と、次機能のテスト実施の開始行の設定準備にもなる）
        For k = 0 To Dic担当者.Count - 1
            If l_str機能 = Dic担当者.Item(Keys_担当者(k))(1) Then
                If Not Dic対象機能.exists(Dic担当者.Item(Keys_担当者(k))(0)) Then
                    Dic対象機能.Add Dic担当者.Item(Keys_担当者(k))(0), Null
                    l_int機能_担当者数 = l_int機能_担当者数 + 1
                End If
            End If
        Next k
        
        'テスト実施者の予実管理表作成(機能単位)
        Call 予実管理明細表作成_機能単位("テスト実施", l_intテスト実施start, Dic対象機能, l_str機能, p_intNissu)

        Dic対象機能.RemoveAll
        
        '################
        '## テスト精査 ##
        '################
        
        'テスト精査の開始行の設定
        '機能単位のテスト実施の開始行 + (日付,担当)(2行分) + 担当者数 + 合計行(1行分) + 空欄(2行分)
        l_intテスト精査start = l_intテスト実施start + 2 + l_int機能_担当者数 + 1 + 2
        
        
        '対象機能の精査者を抽出する。（次の機能に向けて、テスト精査の開始行の設定準備にもなる）
        For k = 0 To Dic精査者.Count - 1
            If l_str機能 = Dic精査者.Item(Keys_精査者(k))(1) Then
                If Not Dic対象機能.exists(Dic精査者.Item(Keys_精査者(k))(0)) Then
                    Dic対象機能.Add Dic精査者.Item(Keys_精査者(k))(0), Null
                    l_int機能_精査者数 = l_int機能_精査者数 + 1
                End If
            End If
        Next k
        
        'テスト精査者の予実管理表作成(機能単位)
        Call 予実管理明細表作成_機能単位("テスト精査", l_intテスト精査start, Dic対象機能, l_str機能, p_intNissu)
        
        Dic対象機能.RemoveAll
        
        '次の機能単位のテスト実施の開始行の設定
        '機能単位のテスト精査開始行 + (日付,担当)(2行分) + 精査者の人数 + 合計行(1行分) + 空白行(4行分)
        l_intテスト実施start = l_intテスト精査start + 2 + l_int機能_精査者数 + 1 + 4
        
        l_int機能_担当者数 = 0
        l_int機能_精査者数 = 0
        l_str機能 = ""
    Next j
    
    '列のグループ化
    For k = 0 To (p_intNissu + 1) * 2
        If (k Mod 14) = 0 And k <> 0 Then
            With Range(Cells(1, k - 13 + 1), Cells(1, k - 1))
                .Columns.Group                                      '非表示状態にしたい個所をグループ化
                ActiveSheet.Outline.ShowLevels columnlevels:=1      '現時点でグループ化されている個所を非表示にする
            End With
        End If
    Next k


    Cells.HorizontalAlignment = xlCenter
    Columns("A").HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter

        
    Application.DisplayAlerts = True    '確認メッセージの表示
    
    Worksheets("日次管理").Activate

End Sub


'###########################
'予実管理明細表作成_機能単位
'###########################
Private Sub 予実管理明細表作成_機能単位(p_strテスト実施_精査 As String, p_int開始行 As Integer, p_dicメンバー As Object, p_str機能 As String, p_intNissu As Integer)
    
    
    '変数宣言
    Dim l_sht_日次管理 As Worksheet
    Dim l_sht_Yojitsu As Worksheet
    Dim Keys() As Variant
    Dim l_intNissu As Integer
    Dim l_wrkDate As Date
    Dim l_intWeeks As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim n As Integer
    Dim l_bolChgFrg As Boolean

    '変数初期化
    Set l_sht_日次管理 = Sheets("日次管理")
    Set l_sht_Yojitsu = Sheets("予実管理_機能単位")
    
    If p_strテスト実施_精査 = "テスト実施" Then     'テスト実施
        With Cells(p_int開始行, 1)
            .Value = p_str機能
            .Font.Bold = True
        End With
        With Cells(p_int開始行, 2)
            .Value = "テスト実施"
            .Font.Bold = True
        End With
    Else                                            'テスト精査
        With Cells(p_int開始行, 1)
            .Value = p_str機能
            .Font.Bold = True
        End With
        With Cells(p_int開始行, 2)
            .Value = "テスト精査"
            .Font.Bold = True
        End With
    End If

    '本日日付
    With Cells(p_int開始行 + 1, 1)
        .Value = "=TODAY()"
        .Font.ColorIndex = 2
        .Font.Bold = True
    End With


    '##############
    'テスト実施者の取得
    '##############
    Cells(p_int開始行 + 2, 1).Value = "担当"
    
    '担当者/精査者をセット
    Keys = p_dicメンバー.Keys
   
    For j = 0 To p_dicメンバー.Count - 1
        Cells(p_int開始行 + 3 + j, 1).Value = Keys(j)
    Next j
    
        
    '合計
    Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1).Value = "合計"

    '１日単位に「予定」「実績」列項目を用意する
    l_intNissu = p_intNissu + 1
    l_wrkDate = l_sht_日次管理.Cells(1, m_intStartColumn).Value
    l_intWeeks = (m_intEndCoumn - m_intStartColumn) / 4 / 7

    For k = 0 To l_intNissu - 1
        For m = 0 To 1
            Cells(p_int開始行 + 1, 2 + (2 * k) + m).Value = l_wrkDate   '基準日

            If ((2 + m) Mod 2) = 0 Then     '予定
                Cells(p_int開始行 + 2, 2 + (2 * k) + m).Value = "予定"
                '担当者
                If p_strテスト実施_精査 = "テスト実施" Then     'テスト実施
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 6).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """実施完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """予定""" + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 4).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行, 1).Address + ")," _
                        + "日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る"
                          
                    Next j
                Else                        'テスト精査
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 11).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """精査完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """予定""" + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 4).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行, 1).Address + ")," _
                        + "日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！"
                    Next j

                End If
                '合計
                Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int開始行 + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + p_dicメンバー.Count - 1, 2 + (2 * k) + m).Address + ")"
                
            ElseIf ((2 + m) Mod 2) = 1 Then  '実績
                Cells(p_int開始行 + 2, 2 + (2 * k) + m).Value = "実績"
                '担当者
                If p_strテスト実施_精査 = "テスト実施" Then     'テスト実施
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 6).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """実施完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """実績""" + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 4).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行, 1).Address + ")," _
                        + "日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！"
                    Next j
                Else                        'テスト精査
                    For j = 0 To p_dicメンバー.Count - 1
                        Cells(p_int開始行 + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((日次管理!" + l_sht_日次管理.Cells(1, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2 + (2 * k) + m).Address + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 11).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1).Address + ")*(日次管理!" + l_sht_日次管理.Cells(2, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(2, m_intEndCoumn).Address + "=" + """精査完了""" + ")*(日次管理!" + l_sht_日次管理.Cells(3, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(3, m_intEndCoumn).Address + "=" + """実績""" + ")*(日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, 4).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int開始行, 1).Address + ")," _
                        + "日次管理!" + l_sht_日次管理.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_日次管理.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '←関数が入る！"
                    Next j
                End If
                '合計
                Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int開始行 + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + p_dicメンバー.Count - 1, 2 + (2 * k) + m).Address + ")"
            End If
        Next m
        l_wrkDate = l_wrkDate + 1
    Next k

    '「実績日による管理」列項目を用意する
    With Range(Cells(p_int開始行 + 1, 1 + l_intNissu * 2 + 1), Cells(p_int開始行 + 1, 1 + l_intNissu * 2 + 2))
        .Merge
        .Value = "実績日による管理" + vbCrLf + "(指摘有無に関わらず" + vbCrLf + "初回の実績入力日で管理)"
        .WrapText = True
        .ColumnWidth = 15
        .RowHeight = 40
    End With
    
    Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 1) = "進捗率"
    '担当者 + 合計
    For j = 0 To p_dicメンバー.Count
        With Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 + 1)
            .Value = "=SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1 + l_intNissu * 2).Address + ")*(" + """実績""" + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2).Address + ")" + _
            "/SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 1, 1 + l_intNissu * 2).Address + ")*(" + """予定""" + "=" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2).Address + ")"
            .NumberFormatLocal = "0%"
        End With
    Next j

    Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2) = "完了率"
    '担当者 + 合計
    For j = 0 To p_dicメンバー.Count
        With Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 + 2)
            .Value = "=" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2).Address + "/" + l_sht_Yojitsu.Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 - 1).Address
            .NumberFormatLocal = "0%"
        End With
    Next j
    
    '背景色
    'Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 40
    'Range(Cells(p_int開始行 + 2, 1), Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 38
    
    For k = 1 To l_intNissu * 2 + 1 Step 2
        With Range(Cells(p_int開始行 + 1, 1 + k), Cells(p_int開始行 + 2, 2 + k))
        
        If l_bolChgFrg Then
            .Interior.ColorIndex = 35
            l_bolChgFrg = False
        Else
            .Interior.ColorIndex = 40
            l_bolChgFrg = True
        End If
        
        End With
    Next k
    
    Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 2, 1)).Interior.ColorIndex = 38
    
        
    '担当者 + 合計
    For j = 0 To p_dicメンバー.Count
        If (j Mod 2) = 1 Then
            Range(Cells(p_int開始行 + 3 + j, 1), Cells(p_int開始行 + 3 + j, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 24
        End If
    Next j

    '罫線描写
    Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1 + l_intNissu * 2 + 2)).Borders.LineStyle = xlContinuous
    
    '太線描写
    '外枠
    With Range(Cells(p_int開始行 + 1, 1), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 1 + l_intNissu * 2 + 2))
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    '内訳
    Range(Cells(p_int開始行 + 2, 1), Cells(p_int開始行 + 2, 1 + l_intNissu * 2 + 2)).Borders(xlEdgeBottom).Weight = xlMedium
    
    For k = 0 To l_intNissu * 2
        If (k Mod 2) = 0 Then
            With Range(Cells(p_int開始行 + 1, 2 + k), Cells(p_int開始行 + 3 + p_dicメンバー.Count, 2 + k))
                .Borders(xlEdgeLeft).Weight = xlMedium
            End With
        End If
    Next k


    Rows(3).Columns.AutoFit                                         'セルの列幅自動設定
    
    For n = 1 To l_intNissu * 2 + 1
        Cells(3, n).ColumnWidth = Cells(3, n).ColumnWidth * 1.2     'セルの列幅自動設定 ×１．２倍
    Next n
    
    '19/03/26 Mod End

End Sub

Private Function 機能取得() As Object

    Dim l_DicKinou As Object
    Dim i As Integer
    Dim l_strKinou As String
    Set l_DicKinou = CreateObject("Scripting.Dictionary")
    
    Worksheets("日次管理").Activate
    
    For i = m_lngStartRow To m_lngEndRow
        '機能を取得
        l_strKinou = Cells(i, 4).Value
        
        If Not l_DicKinou.exists(l_strKinou) Then
            l_DicKinou.Add l_strKinou, Null
        End If
        l_strKinou = ""
    Next i

    Set 機能取得 = l_DicKinou
    
    Worksheets("予実管理_機能単位").Activate
    
End Function

Private Function メンバー取得_機能単位(p_strRoll As String) As Object

    Dim l_Dic As Object
    Dim i As Integer
    Dim j As Integer
    Dim l_strName As String
    Dim l_strKinou As String
    Dim Info(1) As String
    Set l_Dic = CreateObject("Scripting.Dictionary")
    Dim Keys() As Variant
    'l_Dic.RemoveAll
    
    Worksheets("日次管理").Activate
    
    j = 0
    
    For i = m_lngStartRow To m_lngEndRow
        If p_strRoll = "担当者" Then
            '担当者を取得
            l_strName = Cells(i, 6).Value
            l_strKinou = Cells(i, 4).Value
        Else
            '精査者を取得
            l_strName = Cells(i, 11).Value
            l_strKinou = Cells(i, 4).Value
        End If
        If Not l_strName = "" Then
            Info(0) = l_strName
            Info(1) = l_strKinou
            
            l_Dic.Add j, Info

            j = j + 1
        End If
        l_strName = ""
        l_strKinou = ""
        Erase Info
    Next i

    Set メンバー取得_機能単位 = l_Dic

    
    Worksheets("予実管理_機能単位").Activate
    
End Function
'19/04/04 Add End


