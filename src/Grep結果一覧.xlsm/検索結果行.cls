VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "検索結果行"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'（注意１）
'当クラスでは文字列を切り出す際に正規表現を使用する。
'正規表現を使うために、RegExpオブジェクトを利用できるよう「ツール」-「参照設定」で設定している。
'もし、他の端末で使用する場合は、同様の設定を行うか、
'CreateObject関数を使って直接RegExpオブジェクトを取得する。

'（注意２）
'ファイル名は、拡張子vbを検索しているものと想定する。

Private m_lngNo As Long
Private m_str検索文字列  As String
Private m_strフルパス  As String
Private m_strファイル名  As String
Private m_str行数  As String
Private m_str該当箇所  As String

'Input
Public Property Let No(p_lngNo As Long)
    m_lngNo = p_lngNo
End Property
'Output
Public Property Get No() As Long
    No = m_lngNo
End Property
'Input
Public Property Let 検索文字列(p_str検索文字列 As String)
    m_str検索文字列 = p_str検索文字列
End Property
'Output
Public Property Get 検索文字列() As String
    検索文字列 = m_str検索文字列
End Property
'Input
Public Property Let フルパス(p_strフルパス As String)
    m_strフルパス = p_strフルパス
End Property
'Output
Public Property Get フルパス() As String
    フルパス = m_strフルパス
End Property
'Input
Public Property Let ファイル名(p_strファイル名 As String)
    m_strファイル名 = p_strファイル名
End Property
'Output
Public Property Get ファイル名() As String
    ファイル名 = m_strファイル名
End Property
'Input
Public Property Let 行数(p_str行数 As String)
    m_str行数 = p_str行数
End Property
'Output
Public Property Get 行数() As String
    行数 = m_str行数
End Property
'Input
Public Property Let 該当箇所(p_str該当箇所 As String)
    m_str該当箇所 = p_str該当箇所
End Property
'Output
Public Property Get 該当箇所() As String
    該当箇所 = m_str該当箇所
End Property


'文字列の切出し処理が必要な項目について、処理を行ってプロパティに格納する
Public Sub 文字列切出し_データセット(p_str1行 As String)

    Dim l_strWrk文字列 As String
    Dim objRe As RegExp
    Dim objMc As MatchCollection
    Dim l_val文字列抽出 As Variant
    Dim l_str文字列抽出 As String
    Dim l_strTmp As Variant
     
    'オブジェクトセット
    Set objRe = New RegExp
    l_strWrk文字列 = p_str1行
     
     
    '＊＊フルパス＊＊
    '2013/4/4 Ver.1.3修正 start
    'Me.フルパス = l_strWrk文字列
    
    l_strTmp = Split(l_strWrk文字列, "(")
    
    For Each l_val文字列抽出 In l_strTmp
    
        l_str文字列抽出 = CStr(l_val文字列抽出)
    
        '19/05/05 ver2.0 Mod Start
        'objRe.Pattern = "\.vb"
        objRe.Pattern = "\..+$"
        '19/05/05 Mod End
        If objRe.Test(l_str文字列抽出) = True Then
    
            '19/05/05 ver2.0 Mod Start
            'objRe.Pattern = ".+\.vb"
            objRe.Pattern = ".+\..+$"
            '19/05/05 Mod End
            Set objMc = objRe.Execute(l_str文字列抽出)
            
           If objMc.Count > 0 Then
               Me.フルパス = CStr(objMc.Item(0))
               Exit For
           End If
        End If
    Next l_val文字列抽出
    
    l_strTmp = ""
    Set l_val文字列抽出 = Nothing
    l_str文字列抽出 = ""
    Set objMc = Nothing
    '2013/4/4 Ver.1.3修正 end
     
    '＊＊ファイル名＊＊
    l_strTmp = Split(l_strWrk文字列, "\")
     
    For Each l_val文字列抽出 In l_strTmp
     
        l_str文字列抽出 = CStr(l_val文字列抽出)
        
        '19/05/05 ver2.0 Mod Start
        'objRe.Pattern = "\.vb"
        objRe.Pattern = "\..+\([0-9]+.+\[.+?\]: "
        '19/05/05 Mod End
        If objRe.Test(l_str文字列抽出) = True Then
        
            '19/05/05 ver2.0 Mod Start
            'objRe.Pattern = ".+\.vb"
            objRe.Pattern = ".+\..+$"
            objRe.Pattern = "(.+)(\([0-9]+.+\[.+?\]: )(.*)"
            '19/05/05 Mod End
           Set objMc = objRe.Execute(l_str文字列抽出)
           
           If objMc.Count > 0 Then
              '19/05/05 ver2.0 Mod Start
               'Me.ファイル名 = CStr(objMc.Item(0))
               Me.ファイル名 = CStr(objMc(0).SubMatches(0))
               '19/05/05 Mod End
               Exit For
           End If
        End If

    Next l_val文字列抽出
     
    l_strTmp = ""
    Set l_val文字列抽出 = Nothing
    l_str文字列抽出 = ""
    Set objMc = Nothing
    
    
    '＊＊行数＊＊
    l_strTmp = Split(l_strWrk文字列, ",")
    
    For Each l_val文字列抽出 In l_strTmp
    
        l_str文字列抽出 = CStr(l_val文字列抽出)

        '19/05/05 ver2.0 Mod Start
        'objRe.Pattern = "\.vb\("
        objRe.Pattern = "\..+\("
        '19/05/05 Mod End
        If objRe.Test(l_str文字列抽出) = True Then
   
            '19/05/05 ver2.0 Mod Start
            'objRe.Pattern = "(\.vb\()([0-9]+)"
            objRe.Pattern = "(\..+\()([0-9]+)"
            '19/05/05 Mod End
            Set objMc = objRe.Execute(l_str文字列抽出)
    
           If objMc.Count > 0 Then
               Me.行数 = CStr(objMc(0).SubMatches(1))
               Exit For
           End If
        End If
    Next l_val文字列抽出
    
    l_strTmp = ""
    Set l_val文字列抽出 = Nothing
    l_str文字列抽出 = ""
    Set objMc = Nothing
    
    
    '＊＊該当箇所＊＊
    objRe.Pattern = "(\[.+?\]:)"
    If objRe.Test(l_strWrk文字列) = True Then
    
        objRe.Pattern = "(^.+\[.+?\]: )(.+)"

        Set objMc = objRe.Execute(l_strWrk文字列)
    
        If objMc.Count > 0 Then
        '2013/3/26 Ver.1.2修正 start
            Me.該当箇所 = LTrim(CStr(objMc(0).SubMatches(1)))
        '2013/3/26 Ver.1.2修正 end
        '2013/4/9 Ver.1.5修正 start
            If Left(Me.該当箇所, 1) = "=" Then
                Me.該当箇所 = "'" + Me.該当箇所
            End If
        '2013/4/9 Ver.1.5修正 end
        End If
    End If

    l_str文字列抽出 = ""
    Set objMc = Nothing
    
End Sub

