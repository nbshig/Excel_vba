VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 行数指定フォーム 
   Caption         =   "行数指定フォーム"
   ClientHeight    =   2830
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   3630
   OleObjectBlob   =   "行数指定フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "行数指定フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub キャンセル_Click()
    Unload Me
End Sub

Private Sub 作成行確定_Click()

    Dim l_intDec As Integer     '小数点の位置
    Dim l_var行数 As Variant
    l_var行数 = 行数.Value
    l_intDec = InStr(l_var行数, ".")
    
    If Not IsNumeric(l_var行数) Then
        MsgBox "数字を入力して下さい。"
        
    ElseIf l_var行数 <= 0 Then
            MsgBox "1行以上を入力して下さい。"
    ElseIf l_intDec > 0 Then
        MsgBox "整数を入力して下さい。"
        
    Else
        If m_str呼出し元処理 = "初期化" Then
            フォーマット化 (l_var行数)
        ElseIf m_str呼出し元処理 = "行Copy" Then
            行コピー (l_var行数)
        End If
        Unload Me
    End If
    

End Sub
