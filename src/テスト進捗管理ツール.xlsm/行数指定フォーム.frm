VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �s���w��t�H�[�� 
   Caption         =   "�s���w��t�H�[��"
   ClientHeight    =   2830
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   3630
   OleObjectBlob   =   "�s���w��t�H�[��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�s���w��t�H�[��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub �L�����Z��_Click()
    Unload Me
End Sub

Private Sub �쐬�s�m��_Click()

    Dim l_intDec As Integer     '�����_�̈ʒu
    Dim l_var�s�� As Variant
    l_var�s�� = �s��.Value
    l_intDec = InStr(l_var�s��, ".")
    
    If Not IsNumeric(l_var�s��) Then
        MsgBox "��������͂��ĉ������B"
        
    ElseIf l_var�s�� <= 0 Then
            MsgBox "1�s�ȏ����͂��ĉ������B"
    ElseIf l_intDec > 0 Then
        MsgBox "��������͂��ĉ������B"
        
    Else
        If m_str�ďo�������� = "������" Then
            �t�H�[�}�b�g�� (l_var�s��)
        ElseIf m_str�ďo�������� = "�sCopy" Then
            �s�R�s�[ (l_var�s��)
        End If
        Unload Me
    End If
    

End Sub
