Attribute VB_Name = "Module1"
Option Explicit
'################################
'�����o�ϐ��錾
'################################
Dim m_intStartColumn As Integer         '�Ǘ��ΏۂƂȂ�����Ǘ��̊J�n��
Dim m_intEndCoumn As Integer            '�Ǘ��ΏۂƂȂ�����Ǘ��̏I����
Dim m_lngStartRow As Long               '�Ǘ��ΏۂƂȂ�P�[�X�̊J�n�s
Dim m_lngEndRow As Long                 '�Ǘ��ΏۂƂȂ�P�[�X�̏I���s
Dim m_intStartColumnRuiseki As Integer  '�ݐύ��ږ��ׂ̊J�n��
Dim m_intEndCoumnRuiseki As Integer     '�ݐύ��ږ��ׂ̏I����
Public m_str�ďo�������� As String

Sub �����\���Ǘ��ǉ�()

    Application.ScreenUpdating = False  '��ʍX�V�̔�\��

    '##############
    '�����o�ϐ�������
    '##############
    m_intStartColumn = 18   '18���(�J�n��)
    m_intEndCoumn = 0
    m_lngStartRow = 4       '4�s��(�J�n�s)
    m_lngEndRow = Cells(Rows.Count, 2).End(xlUp).Row    '�Ō�̍s
    m_intStartColumnRuiseki = 0
    m_intEndCoumnRuiseki = 0


    '################################
    '�ϐ��錾
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


    '���ʃ`�F�b�N
    If ���ʃ`�F�b�N = False Then
        Exit Sub
    End If


    '�S�̍쐬 or �����쐬�̔��f���s��
    On Error GoTo ErrorHandler_WoekSheetBKUP        '�o�b�N�A�b�v�V�[�g���c���Ă���悤�Ȃ�G���[���b�Z�[�W��\��
    
    If TypeName(Selection) <> "Range" Then
        '�Z���ȊO���I������Ă���Ƃ��Ɏ��s���G���[���������Ȃ��悤�ɏI������
        MsgBox "�Z���ȊO���I������Ă���\��������܂��B" & vbCrLf & "�C�ӂ̃Z����I��������Ԃɂ��Ă����ĉ������B"
        Exit Sub
    Else
        For Each shp In ActiveSheet.CheckBoxes
            With shp
            If .TopLeftCell.Address = "$B$1" Then   '�������ח�̒l�폜���m�F����`�F�b�N�{�b�N�X
                If .Value = 1 Then                 '�`�F�b�N�{�b�N�X��ON�̂Ƃ�
                    l_chkFlg = False
                    .Value = -4146                 '�`�F�b�N�{�b�N�X��������
                Else                               '�`�F�b�N�{�b�N�X��OFF�̂Ƃ�
                    l_chkFlg = True
                    '�o�b�N�A�b�v�V�[�g�쐬
                    Worksheets("�����Ǘ�").Copy after:=Sheets("�����Ǘ�")
                    ActiveSheet.name = "bk_�����Ǘ�"
                    Worksheets("�����Ǘ�").Activate
                    Exit For
                End If
            End If
            End With
        Next
    End If
    
    '##############
    '�N���A����
    '##############
    Call �N���A����
    
    
    '################################
    '�ŏ��̓��t�A�ő�̓��t���擾����
    '################################
    l_arrHiduke = ���t�擾()
    l_datHidukeMin = l_arrHiduke(0) '�ŏ����t
    l_datHidukeMax = l_arrHiduke(1) '�ő���t
    
    
    '�ŏ����t�ƍő���t�̊Ԃ̓������擾����
    l_intNissu = DateDiff("d", l_datHidukeMin, l_datHidukeMax)


    '######################################################################
    '######################################################################
    '�������ח�̍쐬�i�ŏ����t����ő���t�܂ł̓������j
    '######################################################################
    '######################################################################
    
    '##############
    '�^�C�g���ݒ�
    '##############
    l_lngWrkCellY = �^�C�g���ݒ�(l_intNissu, l_datHidukeMin)


    '##############
    '�����o�ϐ��Z�b�g
    '##############
    '�Ǘ��ΏۂƂȂ�����Ǘ��̏I������Z�b�g
    m_intEndCoumn = l_lngWrkCellY - 1
    
    '�ݐύ��ږ��ׂ̊J�n����Z�b�g
    m_intStartColumnRuiseki = m_intEndCoumn + 1
    
    '�ݐύ��ږ��ׂ̏I������Z�b�g
    m_intEndCoumnRuiseki = m_intEndCoumn + 4
    
    
    '##############
    '�ݐύ��ڂ�ݒ肷��
    '##############
    Call �ݐύ��ڒǉ�
    
    
    '###########################
    '�������v/�ݐύ��ڂ�ݒ肷��
    '###########################
    Call �������v_�ݐύ��ڒǉ�



    '##############
    '�r���`��
    '##############
    Call �r���`��(l_intNissu)


    '##############
    '�\���`���ݒ�
    '##############
    Call �\���`���ݒ�


    '##############
    '�t�H�[�}�b�g�p�w�i�F�ݒ�
    '##############
    '�ꎞ�ۊǗp�ϐ�
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
    
    
    Call �t�H�[�}�b�g�p�w�i�F�ݒ�
    
    '���̃����o�ϐ��̒l��߂�
    m_lngStartRow = l_lngWkSR
    m_intStartColumn = l_lngWkSC
    m_intEndCoumn = l_lngWkEC


    '##############
    '�w�i�F�ݒ�
    '##############
    Call �w�i�F�ݒ�(l_intNissu)


    '###########################
    '�`�F�b�N�{�b�N�X�쐬
    '###########################
    Call CheckBox�쐬


    '###########################
    '�`�F�b�N������ꍇ�A�����Ńo�b�N�A�b�v�Ŏ�����l��߂�
    '###########################
    If l_chkFlg Then
        Call �l�Đݒ�(l_datHidukeMin, l_datHidukeMax)
    End If


    '###########################
    '���ѓ���o�^����
    '###########################
    Call ���ѓ��o�^
    
    
    '###########################
    '�o�b�N�A�b�v�V�[�g�̍폜
    '###########################
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.name = "bk_�����Ǘ�" Then
            Application.DisplayAlerts = False
            Worksheets("bk_�����Ǘ�").Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws


    '######################################################################
    '######################################################################
    '�\���Ǘ��V�[�g�̍쐬
    '######################################################################
    '######################################################################
    Call �\���Ǘ��쐬(l_intNissu)
    
    '19/04/04 Add Start
    '######################################################################
    '######################################################################
    '�\���Ǘ�(�@�\�P��)�V�[�g�̍쐬
    '######################################################################
    '######################################################################
    Call �\���Ǘ��쐬_�@�\�P��(l_intNissu)
    '19/04/04 Add End
    
    
    MsgBox "�쐬����"

    Application.ScreenUpdating = True  '��ʍX�V�̕\��
    
    Exit Sub

    '###########################
    '��O����
    '###########################
ErrorHandler_WoekSheetBKUP:
    MsgBox Err.Number & ":" & Err.Description & vbCrLf & "(���́Abk_�����Ǘ��V�[�g���c���Ă���\��������܂��B�폜���ĉ������B)", vbCritical & vbOKOnly, "�G���["

End Sub

Private Function ���t�擾() As Date()

    '################################
    '�ϐ��錾
    '################################
    Dim l_datHidukeArry(2) As Date
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_datwkHiduke  As Date
    Dim i As Integer


    '################################
    '�ϐ�������
    '################################
    l_datHidukeMin = "9999/12/31"
    l_datHidukeMax = "1900/01/01"
    
    
    '################################
    '(�e�X�g���{) ����\���/������ѓ��̒��ōŏ��̓��t���擾����
    '################################
    '(�e�X�g���{) �ŏ��̒���\������擾����
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 7) = "" Then
     Else
      l_datwkHiduke = Cells(i, 7).Value
      If l_datwkHiduke < l_datHidukeMin Then
        l_datHidukeMin = l_datwkHiduke
      End If
     End If
    Next i
    
    '(�e�X�g���{) �ŏ��̒�����ѓ����擾����
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
    '(�e�X�g���{) �擾�����ŏ��̓��t���璼�߂����j���ƂȂ���t���擾����
    '################################
    l_datHidukeMin = searchMonday(l_datHidukeMin)
    
    
    '################################
    '(�e�X�g���{) �����\���/�������ѓ����擾����A�y��
    '(�e�X�g����) �����\���/�������ѓ��̒��ŁA�ő�̓��t���擾����
    '################################
    '(�e�X�g���{) �ő�̊����\������擾����
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 9) = "" Then
     Else
      l_datwkHiduke = Cells(i, 9).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i
    
    '(�e�X�g���{) �ő�̊������ѓ����擾����
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 10) = "" Then
     Else
      l_datwkHiduke = Cells(i, 10).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i
    
    
    '(�e�X�g����) �ő�̊����\������擾����
    For i = m_lngStartRow To m_lngEndRow
     If Cells(i, 12) = "" Then
     Else
      l_datwkHiduke = Cells(i, 12).Value
      If l_datwkHiduke > l_datHidukeMax Then
        l_datHidukeMax = l_datwkHiduke
      End If
     End If
    Next i
    
    '(�e�X�g����) �ő�̊������ѓ����擾����
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
    '(�e�X�g���{) �擾�����ő�̓��t���璼�オ���j���ƂȂ���t���擾����
    '################################
    l_datHidukeMax = searchSunday(l_datHidukeMax)
    
    
    '################################
    '�ŏ����t�A�ő���t��Ԃ�
    '################################
    l_datHidukeArry(0) = l_datHidukeMin
    l_datHidukeArry(1) = l_datHidukeMax

    ���t�擾 = l_datHidukeArry
    
End Function


'################################
'�����̓��t���璼�߂����j���ƂȂ���t��Ԃ�
'################################
Private Function searchMonday(ByVal p_datTmpDate As Date) As Date

    Do While Weekday(p_datTmpDate) <> vbMonday
        p_datTmpDate = p_datTmpDate - 1
    Loop

    searchMonday = p_datTmpDate
    
End Function


'################################
'�����̓��t���璼�オ���j���ƂȂ���t��Ԃ�
'################################
Private Function searchSunday(ByVal p_datTmpDate As Date) As Date

    Do While Weekday(p_datTmpDate) <> vbSunday
        p_datTmpDate = p_datTmpDate + 1
    Loop

    searchSunday = p_datTmpDate
    
End Function


'################################
'�ݐύ��ڂ�ݒ肷��
'################################
Private Sub �ݐύ��ڒǉ�()

    Dim l_lngWrkCellY As Long
    Dim j As Integer
    Dim i As Integer
    Dim l_rngAddress As Range


    '############################
    '�^�C�g���l�Ɗe�ݐϒl�����
    '############################
    '�ϐ�������
    l_lngWrkCellY = 1
    
    For j = m_intStartColumnRuiseki To m_intEndCoumnRuiseki
            If (j Mod 4) = 2 Then                               '���{����-�\��
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "�ݐ�"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "���{����"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "�\��"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=2)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            ElseIf (j Mod 4) = 3 Then                           '���{����-����
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "�ݐ�"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "���{����"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "����"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=3)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            ElseIf (j Mod 4) = 0 Then                           '��������-�\��
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "�ݐ�"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "��������"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "�\��"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=0)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            ElseIf (j Mod 4) = 1 Then                           '��������-����
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "�ݐ�"
                    ElseIf i = 2 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "��������"
                    ElseIf i = 3 Then
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "����"
                    ElseIf Cells(i, 2).Value <> "" Then
                        Set l_rngAddress = Range(Cells(i, m_intStartColumn), Cells(i, m_intEndCoumn))
                        Cells(i, m_intEndCoumn + l_lngWrkCellY) = "=SUMPRODUCT((MOD(COLUMN(" + l_rngAddress.Address(False, True) + "),4)=1)*(" + l_rngAddress.Address(False, True) + "))"
                    End If
                Next i
            Else
                '�����Ȃ�
            End If
            
            l_lngWrkCellY = l_lngWrkCellY + 1
    Next j

End Sub

'################################
'�������v_�ݐύ��ڂ�ݒ肷��
'################################
Private Sub �������v_�ݐύ��ڒǉ�()

    Dim l_lngWrkCellY As Long
    Dim l_rngAddress As Range
    Dim j As Integer
    
    For j = m_intStartColumn To m_intEndCoumnRuiseki
        '�������v�̌v�Z����ݒ�
        Set l_rngAddress = Range(Cells(m_lngStartRow, j), Cells(m_lngEndRow, j))
        Cells(m_lngEndRow + 1, j) = "=SUM(" + l_rngAddress.Address + ")"
        
        '�ݐς̌v�Z����ݒ�
        If j <= m_intStartColumn + 3 Then                       '�ŏ��̓��ɂ��̃P�[�X
            Cells(m_lngEndRow + 2, j) = "=" + Cells(m_lngEndRow + 1, j).Address
        ElseIf j <= m_intEndCoumn Then                          '�ŏ��̓��ɂ�����A�ݐύ��ڂ܂ł̃P�[�X
            Cells(m_lngEndRow + 2, j) = "=" + Cells(m_lngEndRow + 1, j).Address + "+" + Cells(m_lngEndRow + 2, j - 4).Address
        Else
            '�����Ȃ��i�ݐύ��ڂ�z��j
        End If
        
    Next j

End Sub


'################################
'�N���A����
'################################
Public Sub �N���A����()
    Dim shp As CheckBox

    '�������ח�̍폜
    Range("R1", Cells(1, Columns.Count)).EntireColumn.Delete
    
    '���ѓ��̓��͒l�폜
    Range(Cells(m_lngStartRow, 8), Cells(m_lngEndRow, 8)).Value = ""        '�e�X�g���{-����-����
    Range(Cells(m_lngStartRow, 10), Cells(m_lngEndRow, 10)).Value = ""      '�e�X�g���{-����-����
    Range(Cells(m_lngStartRow, 13), Cells(m_lngEndRow, 13)).Value = ""      '�e�X�g����-����-����
    
    '�S�Ẵ`�F�b�N�{�b�N�X���폜����
    For Each shp In ActiveSheet.CheckBoxes
        If shp.TopLeftCell.Address <> "$B$1" Then
                shp.Delete                                                  '�������ח�̒l�폜���m�F����`�F�b�N�{�b�N�X�ȊO�͍폜����
        End If
    Next


End Sub


'################################
'�`�F�b�N�{�b�N�X�쐬
'################################
Sub CheckBox�쐬()

    Dim i As Integer

    'C��ɒl������s�ɂ���A��Ƀ`�F�b�N�{�b�N�X���쐬����
    For i = m_lngStartRow To m_lngEndRow
        Cells(i, 1).Activate
        
        With ActiveSheet.CheckBoxes.Add(1, 1, 1, 1)
            .Height = 10
            .Top = ActiveCell.Top
            .Left = ActiveCell.Left
            .Caption = ""                           '�e�L�X�g
            .Value = False
        End With
     Next i
     
End Sub


'################################
'�s�폜
'�`�F�b�N�{�b�N�X��ON�̍s���폜����
'################################
Sub �s�폜()

    Dim i As Integer
    Dim shp As CheckBox
    Dim rng As Range
    Dim s As String
    
    '##############
    '�����o�ϐ�������
    '##############
    m_intStartColumn = 18
    m_intEndCoumn = 0
    m_lngStartRow = 4
    m_lngEndRow = Cells(Rows.Count, 2).End(xlUp).Row
    m_lngEndRow = Range("B" & Rows.Count).End(xlUp).Row '������
    m_intStartColumnRuiseki = 0
    m_intEndCoumnRuiseki = 0


    If TypeName(Selection) <> "Range" Then
        '�Z���ȊO���I������Ă���Ƃ��Ɏ��s���G���[���������Ȃ��悤�ɏI������
        MsgBox "�Z���ȊO���I������Ă���\��������܂��B" & vbCrLf & "�C�ӂ̃Z����I��������Ԃɂ��Ă����ĉ������B"
        Exit Sub
    Else
      For Each shp In ActiveSheet.CheckBoxes
        If shp.Value = 1 Then
            '�`�F�b�N�{�b�N�XON�̂Ƃ�
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
'�sCopy
'���ׂ̍ŏI�s���烌�C�A�E�g���R�s�[����
'################################
Sub �sCopy()
    m_str�ďo�������� = "�sCopy"
    �s���w��t�H�[��.Show

End Sub


'################################
'�sCopy
'�s���w��t�H�[�����畜�A
'################################
Sub �s�R�s�[(p_int�s�� As Integer)
    Application.ScreenUpdating = False  '��ʍX�V�̔�\��
    
    Dim i As Integer
    Dim s As String
    ReDim l_arrHiduke(2) As Date
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_intNissu As Integer

    '##############
    '�����o�ϐ�������
    '##############
    m_intStartColumn = 18
    m_lngStartRow = 4
    m_lngEndRow = Cells(Rows.Count, 3).End(xlUp).Row
    m_lngEndRow = Range("B" & Rows.Count).End(xlUp).Row '������

    '���ʃ`�F�b�N�i���݂̍ŏI�s�m�F�܂ށj
    If ���ʃ`�F�b�N = False Then
        Exit Sub
    End If
    
    '���ʃ`�F�b�N��A���߂ă����o�ϐ�������
    'm_intEndCoumnRuiseki = Cells(m_lngEndRow, Columns.Count).End(xlToLeft).Column
    m_intEndCoumnRuiseki = Cells(1, Columns.Count).End(xlToLeft).Column '�������ח�̗L���m�F
    
    m_intStartColumnRuiseki = m_intEndCoumnRuiseki - 3
    m_intEndCoumn = m_intEndCoumnRuiseki - 4

    '�`�F�b�N�i�������ח�̗L���j
    If m_intStartColumnRuiseki <= 17 Then
        MsgBox "�������ח񂪂܂��쐬����Ă��Ȃ��悤�ł��B" & vbCrLf & "��ɐi���Ǘ��{�^�������s���ĉ������B"
        Exit Sub
    End If
    
    Range(Cells(m_lngEndRow, 1), Cells(m_lngEndRow, 1)).EntireRow.Copy
    Range(Cells(m_lngEndRow + 1, 1), Cells(m_lngEndRow + p_int�s��, 1)).EntireRow.PasteSpecial
    Range(Cells(m_lngEndRow + 1, 1), Cells(m_lngEndRow + p_int�s��, m_intEndCoumn)).ClearContents 'C�񂪃u�����N���Ɩ��׍s(���^�C�g��)�̔w�i�F��ݒ肵�Ă��܂��̂ŁA"�s�R�s�["���Z�b�g�AD��ȍ~�̒l���N���A
    Range(Cells(m_lngEndRow + 1, 3), Cells(m_lngEndRow + p_int�s��, 3)).Value = "�s�R�s�["
    
    m_lngEndRow = m_lngEndRow + p_int�s��


    '###########################
    '�������v/�ݐύ��ڂ�ݒ肷��
    '###########################
    Call �������v_�ݐύ��ڒǉ�
    
    
    '################################
    '�ŏ��̓��t�A�ő�̓��t���擾����
    '################################
    l_arrHiduke = ���t�擾()
    l_datHidukeMin = l_arrHiduke(0) '�ŏ����t
    l_datHidukeMax = l_arrHiduke(1) '�ő���t
    
    '�ŏ����t�ƍő���t�̊Ԃ̓������擾����
    l_intNissu = DateDiff("d", l_datHidukeMin, l_datHidukeMax)


    '##############
    '�t�H�[�}�b�g�p�w�i�F�ݒ�
    '##############
    '�ꎞ�ۊǗp�ϐ�
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
    
    Call �t�H�[�}�b�g�p�w�i�F�ݒ�
    
    '���̃����o�ϐ��̒l��߂�
    m_lngStartRow = l_lngWkSR
    m_intStartColumn = l_lngWkSC
    m_intEndCoumn = l_lngWkEC
    
    
    '##############
    '�r���`��
    '##############
    Call �r���`��(l_intNissu)
    

    '##############
    '�\���`���ݒ�
    '##############
    Call �\���`���ݒ�


    '##############
    '�w�i�F�ݒ�
    '##############
    Call �w�i�F�ݒ�(l_intNissu)


    '##############
    '�f�[�^�̓��͋K���ݒ�
    '##############
    Call �f�[�^���͋K���ݒ�
    
    
    '###########################
    '�`�F�b�N�{�b�N�X�쐬
    '###########################
    m_lngStartRow = m_lngEndRow - p_int�s�� + 1   '�R�s�[�s�����`�F�b�N�{�b�N�X���쐬����
    
    Call CheckBox�쐬

    Application.ScreenUpdating = True  '��ʍX�V�̕\��
End Sub


'##############
'�r���`��
'##############
Private Sub �r���`��(p_intNissu As Integer)

    Dim l_lngWrkCellY As Long
    Dim j As Integer
    Dim l_intNissu As Integer
    
    '�ϐ�������
    l_lngWrkCellY = m_intStartColumn
    l_intNissu = p_intNissu
    
    '�r���`��
    Range(Cells(2, m_intStartColumn), Cells(m_lngEndRow + 2, m_intEndCoumnRuiseki)).Borders.LineStyle = xlContinuous
    
    '�����`��
    For j = m_intStartColumn To l_intNissu + m_intStartColumn + 1   '"+1"�͗ݐύ��ڕ�
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
'�\���`���ݒ�
'##############
Private Sub �\���`���ݒ�()

    '���t�`���𐮂���
    Range(Cells(1, m_intStartColumn), Cells(1, m_intEndCoumn)).NumberFormatLocal = "mm��dd��"
    
    '�����ɂ���
    Range(Cells(m_lngStartRow, m_intStartColumnRuiseki), Cells(m_lngEndRow + 1, m_intEndCoumnRuiseki)).Font.Bold = True
    
End Sub


'##############
'�������ח�y�у^�C�g���s�̔w�i�F�ݒ�
'##############
Private Sub �w�i�F�ݒ�(p_intNissu As Integer)

    Dim j As Integer
    Dim k As Integer
    Dim i As Integer
    Dim l_intNissu As Integer
    Dim l_lngWrkCellY As Long
    Dim l_lngWrkCellYWeek As Long

    '�ϐ�������
    l_lngWrkCellY = m_intStartColumn
    l_lngWrkCellYWeek = m_intStartColumn
    l_intNissu = p_intNissu
    
    For j = m_intStartColumn To l_intNissu + m_intStartColumn + 1     '��P�ʂ̏����B ""+1"�͗ݐύ��ڕ�
        '���׍s�̔w�i�F�ݒ�
        For k = 0 To 3                  '1��4�񕪂̏���
            'If ((l_lngWrkCellY + k) Mod 2) = 1 Then
            If ((l_lngWrkCellY + k) Mod 4) = 2 Then                   '���{����-�\��
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 40
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 3 Then               '���{����-����
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 24
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 0 Then               '��������-�\��
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 35
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 1 Then               '��������-����
                Range(Cells(m_lngStartRow, l_lngWrkCellY + k), Cells(m_lngEndRow, l_lngWrkCellY + k)).Interior.ColorIndex = 20
            Else
                '�����Ȃ�
            End If
        Next k
        
        '�^�C�g���s�̔w�i�F�ݒ�
        If l_lngWrkCellYWeek < m_intEndCoumn Then
            If (j Mod 2) = 0 Then '1�T�ԕ��i28��)���̏���
                Range(Cells(1, l_lngWrkCellYWeek), Cells(3, l_lngWrkCellYWeek + 27)).Interior.ColorIndex = 42
            Else
                Range(Cells(1, l_lngWrkCellYWeek), Cells(3, l_lngWrkCellYWeek + 27)).Interior.ColorIndex = 34
            End If
        End If
        l_lngWrkCellY = l_lngWrkCellY + 4
        l_lngWrkCellYWeek = l_lngWrkCellYWeek + 28
    Next j
    
    '���׍s(���^�C�g��)�̔w�i�F�ݒ�
    For i = 4 To m_lngEndRow       '�s�P�ʂ̏���
            If Cells(i, 2).Value = "" And Cells(i, 5).Value = "" And Cells(i, 7).Value = "" And Cells(i, 9).Value = "" And Cells(i, 12).Value = "" Then
                Range(Cells(i, 1), Cells(i, m_intEndCoumn + 4)).Interior.ColorIndex = 33
            End If
    Next i
    
    '�ݐύ��ڃ^�C�g���s�̔w�i�F�ݒ�
    Range(Cells(1, m_intStartColumnRuiseki), Cells(3, m_intEndCoumnRuiseki)).Interior.ColorIndex = 44
    
    '�ݐύ��ڂ̍��v�s�̔w�i�F�ݒ�
    Range(Cells(m_lngEndRow + 1, m_intStartColumnRuiseki), Cells(m_lngEndRow + 1, m_intEndCoumnRuiseki)).Interior.ColorIndex = 27

End Sub


'##############
'�^�C�g���ݒ�
'##############
Private Function �^�C�g���ݒ�(p_intNissu As Integer, p_datHidukeMin As Date) As Long

    Dim j As Integer
    Dim k As Integer
    Dim i As Integer
    Dim l_lngWrkCellY As Long
    Dim l_intNissu As Integer
    Dim l_datHidukeMin As Date

    
    '�ϐ�������
    l_lngWrkCellY = m_intStartColumn
    l_intNissu = p_intNissu
    l_datHidukeMin = p_datHidukeMin
    
    For j = m_intStartColumn To l_intNissu + m_intStartColumn   '��P�ʂ̏���
        For k = 0 To 3                                          '1��4�񕪂̏���
            If ((l_lngWrkCellY + k) Mod 4) = 2 Then             '���{����-�\��
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "���{����"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "�\��"
                    End If
                Next i
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 3 Then         '���{����-����
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "���{����"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "����"
                    End If
                Next i
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 0 Then         '��������-�\��
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "��������"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "�\��"
                    End If
                Next i
            ElseIf ((l_lngWrkCellY + k) Mod 4) = 1 Then         '��������-����
                For i = 1 To m_lngEndRow                        '�s�P�ʂ̏���
                    If i = 1 Then
                        Cells(i, l_lngWrkCellY + k) = l_datHidukeMin + (j - m_intStartColumn)
                    ElseIf i = 2 Then
                        Cells(i, l_lngWrkCellY + k) = "��������"
                    ElseIf i = 3 Then
                        Cells(i, l_lngWrkCellY + k) = "����"
                    End If
                Next i
            Else
                '�����Ȃ�
            End If
        Next k
        
        l_lngWrkCellY = l_lngWrkCellY + 4
    Next j
    
    �^�C�g���ݒ� = l_lngWrkCellY
    
End Function


Sub ������()

    m_str�ďo�������� = "������"
    �s���w��t�H�[��.Show

End Sub
 
Public Sub �t�H�[�}�b�g��(p_int�s�� As Integer)
    Application.ScreenUpdating = False  '��ʍX�V�̔�\��
    
    Dim i As Integer
    Dim s As String
    ReDim l_arrHiduke(2) As Date
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_intNissu As Integer
    

    '##############
    '�����o�ϐ�������
    '##############
    m_intStartColumn = 1
    m_intEndCoumn = 17
    m_lngStartRow = 4
    m_lngEndRow = 4 + p_int�s�� - 1


    '################################
    '�t�H�[�}�b�g�p�N���A����
    '################################
    �t�H�[�}�b�g�p�N���A����


    '##############
    '�r���`��
    '##############
    Call �t�H�[�}�b�g�p�r���`��
    

    '##############
    '�\���`���ݒ�
    '##############
    Call �t�H�[�}�b�g�p�\���`���ݒ�


    '##############
    '�t�H�[�}�b�g�p�w�i�F�ݒ�
    '##############
    Call �t�H�[�}�b�g�p�w�i�F�ݒ�
    
    
    '##############
    '�f�[�^�̓��͋K���ݒ�
    '##############
    Call �f�[�^���͋K���ݒ�

    
    '###########################
    '�`�F�b�N�{�b�N�X�쐬
    '###########################
    Call CheckBox�쐬
    
    MsgBox "����������"

    Application.ScreenUpdating = True  '��ʍX�V�̕\��
End Sub

'#######################
'�t�H�[�}�b�g�p�r���`��
'#######################
Private Sub �t�H�[�}�b�g�p�r���`��()
    
    '�r���`��
    Range(Cells(m_lngStartRow, m_intStartColumn), Cells(m_lngEndRow, m_intEndCoumn)).Borders.LineStyle = xlContinuous
    
End Sub

'###########################
'�t�H�[�}�b�g�p�\���`���ݒ�
'###########################
Private Sub �t�H�[�}�b�g�p�\���`���ݒ�()

    '���t�`���𐮂���
    Range(Cells(m_lngStartRow, 7), Cells(m_lngEndRow, 10)).NumberFormatLocal = "mm��dd��"
    Range(Cells(m_lngStartRow, 12), Cells(m_lngEndRow, 13)).NumberFormatLocal = "mm��dd��"
    Range(Cells(m_lngStartRow, 15), Cells(m_lngEndRow, 16)).NumberFormatLocal = "mm��dd��"

    '�z�u�𐮂���
    Cells.HorizontalAlignment = xlCenter
    Range(Cells(m_lngStartRow, 3), Cells(m_lngEndRow, 3)).HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter
    
End Sub


'##########################
'�t�H�[�}�b�g�p�w�i�F�ݒ�
'##########################
Private Sub �t�H�[�}�b�g�p�w�i�F�ݒ�()
    
    Range(Cells(m_lngStartRow, m_intStartColumn), Cells(m_lngEndRow, m_intEndCoumn)).Interior.ColorIndex = 2
    
    '���̓Z���ɂ��Ĕw�i�F��ݒ肷��
    Range(Cells(m_lngStartRow, 2), Cells(m_lngEndRow, 2)).Interior.ColorIndex = 36       'ID
    Range(Cells(m_lngStartRow, 3), Cells(m_lngEndRow, 3)).Interior.ColorIndex = 19      '�V�i���I��
    '19/04/04 Add Start
    Range(Cells(m_lngStartRow, 4), Cells(m_lngEndRow, 3)).Interior.ColorIndex = 19      '�@�\��
    '19/04/04 Add End
    Range(Cells(m_lngStartRow, 5), Cells(m_lngEndRow, 5)).Interior.ColorIndex = 19      '�P�[�X��
    Range(Cells(m_lngStartRow, 6), Cells(m_lngEndRow, 6)).Interior.ColorIndex = 19      '���{��
    Range(Cells(m_lngStartRow, 7), Cells(m_lngEndRow, 7)).Interior.ColorIndex = 36       '�e�X�g���{-����-�\��
    Range(Cells(m_lngStartRow, 9), Cells(m_lngEndRow, 9)).Interior.ColorIndex = 36       '�e�X�g���{-����-�\��
    Range(Cells(m_lngStartRow, 11), Cells(m_lngEndRow, 11)).Interior.ColorIndex = 19    '������
    Range(Cells(m_lngStartRow, 12), Cells(m_lngEndRow, 12)).Interior.ColorIndex = 36     '�e�X�g����-����-�\��
    Range(Cells(m_lngStartRow, 14), Cells(m_lngEndRow, 14)).Interior.ColorIndex = 19    '�w�E�L��
    Range(Cells(m_lngStartRow, 15), Cells(m_lngEndRow, 15)).Interior.ColorIndex = 19    '�w�E�Ή�-(���{��)�w�E�C����
    Range(Cells(m_lngStartRow, 16), Cells(m_lngEndRow, 16)).Interior.ColorIndex = 19    '�w�E�Ή�-(������)�w�E�m�F��
    Range(Cells(m_lngStartRow, 17), Cells(m_lngEndRow, 17)).Interior.ColorIndex = 19    '��
End Sub


'##########################
'�f�[�^���͋K���ݒ�
'##########################
Private Sub �f�[�^���͋K���ݒ�()

    Range(Cells(m_lngStartRow, 14), Cells(m_lngEndRow, 14)).Select
    With Range(Cells(m_lngStartRow, 14), Cells(m_lngEndRow, 14)).Validation
            .Delete
            .Add Type:=xlValidateList, _
                Formula1:="�L,��"
    End With
    
    Range(Cells(m_lngStartRow, 17), Cells(m_lngEndRow, 17)).Select
    With Range(Cells(m_lngStartRow, 17), Cells(m_lngEndRow, 17)).Validation
            .Delete
            .Add Type:=xlValidateList, _
                Formula1:="���{��,������,�Ď��{��,�Đ�����,����"
    End With

End Sub


'################################
'�t�H�[�}�b�g�p�N���A����
'################################
Public Sub �t�H�[�}�b�g�p�N���A����()

    Dim shp As CheckBox

    Range("R1", Cells(1, Columns.Count)).EntireColumn.Delete
    Range("A4", Cells(Rows.Count, 2)).EntireRow.Delete
    
    '�`�F�b�N�{�b�N�X���폜����
    'ActiveSheet.CheckBoxes.Delete
    For Each shp In ActiveSheet.CheckBoxes
        If shp.TopLeftCell.Address <> "$B$1" Then
                shp.Delete                          '�������ח�̒l�폜���m�F����`�F�b�N�{�b�N�X�ȊO�͍폜����
        End If
    Next

End Sub


'################################
'���ʃ`�F�b�N
'################################
Public Function ���ʃ`�F�b�N() As Boolean

    '�ϐ��錾
    Dim l_datHidukeMin As Date
    Dim l_datHidukeMax As Date
    Dim l_datwkHiduke  As Date
    Dim i As Integer
    
    '�ϐ�������
    ���ʃ`�F�b�N = True
    l_datHidukeMin = "9999/12/31"
    l_datHidukeMax = "1900/01/01"
    
    
    '<�`�F�b�N>
    '�@�R�s�[�Ώۍs�̑��݃`�F�b�N
    If m_lngEndRow <= 3 Then
        MsgBox "4�s�ڈȍ~��ID����͂��ĉ������B"
        
        ���ʃ`�F�b�N = False
        Exit Function
    End If
    
    
    '�A �\����̓��̓`�F�b�N
    '(�e�X�g���{) �ŏ��̒���\������擾����
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
        MsgBox "�e�X�g���{-����\������Œ�P�͓��͂��ĉ������B"
        
        ���ʃ`�F�b�N = False
        Exit Function
    End If
    
    '(�e�X�g���{) �ő�̊����\������擾����
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
        MsgBox "�e�X�g���{-�����\������Œ�P�͓��͂��ĉ������B"
        
        ���ʃ`�F�b�N = False
        Exit Function
    End If
    
    '(�e�X�g����) �ő�̊����\������擾����
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
        MsgBox "�e�X�g����-�����\������Œ�P�͓��͂��ĉ������B"
        
        ���ʃ`�F�b�N = False
        Exit Function
    End If
    
End Function


'�o�b�N�A�b�v�����V�[�g����l���Đݒ肷��
'################################
Private Sub �l�Đݒ�(p_datHidukeMin As Date, p_datHidukeMax As Date)


    Worksheets("bk_�����Ǘ�").Activate
    
    Dim l_intLastColumn_bk�����Ǘ� As Integer
    Dim l_datHidukeMin_bk�����Ǘ� As Date
    Dim l_datHidukeMax_bk�����Ǘ� As Date
    Dim l_intNissu_bk�����Ǘ� As Integer
        
    'bk_�����Ǘ��V�[�g��̃Z��R1�̍ŏ����t���擾
    l_datHidukeMin_bk�����Ǘ� = Cells(1, 18).Value
    
    'bk_�����Ǘ��V�[�g��̍ő���t���擾
    l_intLastColumn_bk�����Ǘ� = Cells(1, Columns.Count).End(xlToLeft).Column       '�ŏI��
    l_datHidukeMax_bk�����Ǘ� = Cells(1, l_intLastColumn_bk�����Ǘ� - 4)            '�ŏI�񂩂�ݐ�4���ڂ�����
    
    '�����Ǘ��V�[�g�̍ŏ����t��bk_�����Ǘ��V�[�g��̍ŏ����t���r����
    If p_datHidukeMin < l_datHidukeMin_bk�����Ǘ� Then
        '�����Ǘ��V�[�g�̍ŏ����t�̕����������Ƃ��Abk_�����Ǘ��V�[�g�̍ŏ����t����l���擾���Ă��A�����Ǘ��V�[�g��œ\��t����Z���͑��݂���
        '����āAbk_�����Ǘ��V�[�g��̍ŏ����t��ݒ肷��
        l_datHidukeMin_bk�����Ǘ� = l_datHidukeMin_bk�����Ǘ�                       '���ɏ����ɈӖ��Ȃ�
    Else
        'bk_�����Ǘ��V�[�g�̍ŏ����t�̕����������Ƃ��Abk_�����Ǘ��V�[�g�̍ŏ����t����l���擾����ƁA�����Ǘ��V�[�g��œ\��t����Z���͑��݂��Ȃ����ƂɂȂ�
        '����āA�����Ǘ��V�[�g��̍ŏ����t��ݒ肷��
        l_datHidukeMin_bk�����Ǘ� = p_datHidukeMin
    End If
    
    '�����Ǘ��V�[�g�̍ő���t��bk_�����Ǘ��V�[�g��̍ő���t���r����
    If p_datHidukeMax > l_datHidukeMax_bk�����Ǘ� Then
        '�����Ǘ��V�[�g�̍ő���t�̕����傫���Ƃ��Abk_�����Ǘ��V�[�g�̍ő���t����l���擾���Ă��A�����Ǘ��V�[�g��œ\��t����Z���͑��݂���
        '����āAbk_�����Ǘ��V�[�g��̍ő���t��ݒ肷��
        If l_datHidukeMax_bk�����Ǘ� = "0:00:00" Then                               '�A���A����̐i���Ǘ��{�^����������bk_�����Ǘ��V�[�g��ɓ������ח񂪂Ȃ��̂ŁA���t�͓����Ǘ��V�[�g�ォ��擾����
            l_datHidukeMax_bk�����Ǘ� = p_datHidukeMax
        Else
            l_datHidukeMax_bk�����Ǘ� = l_datHidukeMax_bk�����Ǘ�                   '���ɏ����ɈӖ��Ȃ�
        End If
    Else
        'bk_�����Ǘ��V�[�g�̍ő���t�̕����傫���Ƃ��Abk_�����Ǘ��V�[�g�̍ő���t����l���擾����ƁA�����Ǘ��V�[�g��œ\��t����Z���͑��݂��Ȃ����ƂɂȂ�
        '����āA�����Ǘ��V�[�g��̍ő���t��ݒ肷��
        l_datHidukeMax_bk�����Ǘ� = p_datHidukeMax
    End If
    
    '�ŏ����t�ƍő���t�̊Ԃ̓������擾����
    l_intNissu_bk�����Ǘ� = DateDiff("d", l_datHidukeMin_bk�����Ǘ�, l_datHidukeMax_bk�����Ǘ�)     '(��)20���`26���̃J�E���g��6���ԁB�i�܂�J�n��(20��)�̓J�E���g���Ă��Ȃ��j
    l_intNissu_bk�����Ǘ� = l_intNissu_bk�����Ǘ� + 1                                               '���̂��߁A�{1������


    '############################
    'bk_�����Ǘ��V�[�g��̊Ǘ����ׂ̒l������Ǘ��V�[�g��̓������t�̍��ڂɃZ�b�g����
    '############################
    '�����Ǘ��V�[�g��ŏ�L�Ō������ŏ����t���Z�b�g���Ă���񍀖ڂ�T������
    Worksheets("�����Ǘ�").Activate
    
    Dim j As Integer
    Dim l_intLastColumn_�����Ǘ� As Integer
    Dim l_datHidukeMin_�����Ǘ� As Date
    Dim l_datHidukeMax_�����Ǘ� As Date
    Dim l_intNissu_�����Ǘ� As Integer
    
    l_datHidukeMin_�����Ǘ� = Cells(1, 18).Value
    l_intLastColumn_�����Ǘ� = Cells(1, Columns.Count).End(xlToLeft).Column '�ŏI��
    l_datHidukeMax_�����Ǘ� = Cells(1, l_intLastColumn_�����Ǘ� - 4)                        '�ŏI�񂩂�ݐ�4���ڂ�����
    
    '�ŏ����t�ƍő���t�̊Ԃ̓������擾����
    l_intNissu_�����Ǘ� = DateDiff("d", l_datHidukeMin_�����Ǘ�, l_datHidukeMax_�����Ǘ�)   '(��)20���`26���̃J�E���g��6���ԁB�i�܂�J�n��(20��)�̓J�E���g���Ă��Ȃ��j
    l_intNissu_�����Ǘ� = l_intNissu_�����Ǘ� + 1                                           '���̂��߁A�{1������
    
    j = 0
    
    For j = m_intStartColumn To l_intNissu_�����Ǘ� * 4 + m_intStartColumn + 1              '��P�ʂ̏����B ""+1"�͗ݐύ��ڕ�
        'bk_�����Ǘ��V�[�g��̍ŏ����t�Ɠ����Ǘ��V�[�g��̓��t����v����Z�����m�F
        If l_datHidukeMin_bk�����Ǘ� = Range(Cells(1, j), Cells(1, j)).Value Then
            '�����Ǘ��V�[�g��̍ŏ����t�����������ꍇ�Abk_�����Ǘ��V�[�g�ォ��l���Z���͈͈ꊇ�ŃR�s�[���y�[�X�g
            Worksheets("bk_�����Ǘ�").Activate
            Worksheets("bk_�����Ǘ�").Range(Cells(m_lngStartRow, m_intStartColumn), Cells(m_lngEndRow, m_intStartColumn - 1 + (l_intNissu_bk�����Ǘ� * 4))).Copy
                       
            Worksheets("�����Ǘ�").Activate
            ActiveSheet.Range(ActiveSheet.Cells(m_lngStartRow, j), ActiveSheet.Cells(m_lngEndRow, j - 1 + (l_intNissu_bk�����Ǘ� * 4))).PasteSpecial Paste:=xlPasteValues
            
            Exit For
        End If
    Next j

End Sub


'################################
'�\���Ǘ��V�[�g���쐬����
'################################
Private Sub �\���Ǘ��쐬(p_intNissu As Integer)

    Worksheets("�\���Ǘ�").Activate

    '�V�[�g������
    Application.DisplayAlerts = False   '�m�F���b�Z�[�W�̔�\��
    Range("3:3").UnMerge                '�Z���̌�������
    Cells.ClearFormats                  '�S�Z���̕\���`���N���A
    Cells.Clear                         '�S�Z���̃N���A
    Cells.ColumnWidth = 5               '�S�Z���̕�
    Cells.RowHeight = 20                '�S�Z���̍���
    '19/03/26 Add Start
    ActiveSheet.Columns.ClearOutline    '��̃O���[�v����S�č폜
    '19/03/26 Add End

    '�ϐ��錾
    Dim Dic�S���� As Object
    Dim Dic������ As Object
    Dim l_int�e�X�g���{start As Integer
    Dim l_int�e�X�g����start As Integer
    Dim i As Integer
    
    '�^�C�g��
    With Cells(1, 1)
        .Value = "�\���Ǘ�"
        .Font.Bold = True
    End With
    '19/03/26 DEL Start
'    With Cells(1, 2)
'        .Value = "(���j�����T����Ƃ���)"
'        .Font.Bold = True
'    End With
    '19/03/26 DEL End

    '###########################
    '## �e�X�g���{ ##
    '###########################
    '�e�X�g���{�̊J�n�s�̐ݒ�
    l_int�e�X�g���{start = 2
    
    '�e�X�g���{�҂̎擾
    Set Dic�S���� = �����o�[�擾("�S����")
    
    '�e�X�g���{�҂̗\���Ǘ��\�쐬
    Call �\���Ǘ����ו\�쐬(l_int�e�X�g���{start, Dic�S����, p_intNissu)


    '###########################
    '## �e�X�g���� ##
    '###########################
    '�e�X�g�����̊J�n�s�̐ݒ�
    l_int�e�X�g����start = 5 + Dic�S����.Count + 3

    '�e�X�g�����҂̎擾
    Set Dic������ = �����o�[�擾("������")

    '�e�X�g�����҂̗\���Ǘ��\�쐬
    Call �\���Ǘ����ו\�쐬(l_int�e�X�g����start, Dic������, p_intNissu)

    Cells.HorizontalAlignment = xlCenter
    Cells(1, 2).HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter

        
    Application.DisplayAlerts = True    '�m�F���b�Z�[�W�̕\��
    
    Worksheets("�����Ǘ�").Activate

End Sub


Private Function �����o�[�擾(p_strRoll As String) As Object

    Dim Dic As Object
    Dim i As Integer
    Dim l_strName As String
    Set Dic = CreateObject("Scripting.Dictionary")
    
    Worksheets("�����Ǘ�").Activate
    
    For i = m_lngStartRow To m_lngEndRow
        If p_strRoll = "�S����" Then
            '�S���҂��擾
            l_strName = Cells(i, 6).Value
        Else
            '�����҂��擾
            l_strName = Cells(i, 11).Value
        End If
        If Not l_strName = "" Then
            If Not Dic.exists(l_strName) Then
                Dic.Add l_strName, Null
            End If
        End If
    Next i

    Set �����o�[�擾 = Dic
    
    Worksheets("�\���Ǘ�").Activate
    
End Function


'###########################
'���ѓ���o�^����
'###########################
Private Sub ���ѓ��o�^()

    '�ϐ��錾
    Dim i As Integer
    Dim j As Integer
    
    For i = m_lngStartRow To m_lngEndRow
    
        '19/03/26 Add Start
        '�e�X�g���{-����-���ђl�̓o�^
        For j = m_intStartColumn To m_intEndCoumn Step 1        '�������ח�������ɌJ���
            If (j Mod 4) = 3 Then                               '���{����-���т̗�
                If Cells(i, j).Value <> "" Then
                    Cells(i, 8).Value = Cells(1, j).Value
                    Exit For
                End If
            End If
        Next j
        '19/03/26 Add End
    
        '�ݐώ��{����-�\�萔/���ѐ��̈�v���m�F
        If Cells(i, m_intStartColumnRuiseki).Value = Cells(i, m_intStartColumnRuiseki + 1).Value Then
            '�e�X�g���{-����-���ђl�̓o�^
            For j = m_intEndCoumn To m_intStartColumn Step -1       '�������ח���~���ɌJ���
                If (j Mod 4) = 3 Then                               '���{����-���т̗�
                    If Cells(i, j).Value <> "" Then
                        Cells(i, 10).Value = Cells(1, j).Value
                        Exit For
                    End If
                End If
            Next j

            '�e�X�g���{-����-���ђl�̓o�^
            For j = m_intStartColumn To m_intEndCoumn Step 1        '�������ח�������ɌJ���
                If (j Mod 4) = 3 Then                               '���{����-���т̗�
                    If Cells(i, j).Value <> "" Then
                        Cells(i, 8).Value = Cells(1, j).Value
                        Exit For
                    End If
                End If
            Next j
                
        End If
        
        '�ݐϐ�������-�\�萔/���ѐ��̈�v���m�F
        If Cells(i, m_intStartColumnRuiseki + 2).Value = Cells(i, m_intStartColumnRuiseki + 3).Value Then
            '�e�X�g����-����-���ђl�̓o�^
            For j = m_intEndCoumn To m_intStartColumn Step -1       '�������ח���~���ɌJ���
                If (j Mod 4) = 1 Then                               '��������-���т̗�
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
'�\���Ǘ����ו\�쐬
'###########################
Private Sub �\���Ǘ����ו\�쐬(p_int�J�n�s As Integer, p_dic�����o�[ As Object, p_intNissu As Integer)

    '�ϐ��錾
    Dim l_sht_�����Ǘ� As Worksheet
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

    '�ϐ�������
    Set l_sht_�����Ǘ� = Sheets("�����Ǘ�")
    Set l_sht_Yojitsu = Sheets("�\���Ǘ�")
    
    If p_int�J�n�s = 2 Then     '�e�X�g���{
        With Cells(p_int�J�n�s, 1)
            .Value = "�e�X�g���{"
            .Font.Bold = True
        End With
        With Cells(p_int�J�n�s, 2)
            .Value = "(�ݐϊǗ�)"
            .Font.Bold = True
        End With
    Else                        '�e�X�g����
        With Cells(p_int�J�n�s, 1)
            .Value = "�e�X�g����"
            .Font.Bold = True
        End With
        With Cells(p_int�J�n�s, 2)
            .Value = "(�ݐϊǗ�)"
            .Font.Bold = True
        End With
    End If

    '�{�����t
    With Cells(p_int�J�n�s + 1, 1)
        .Value = "=TODAY()"
        .Font.ColorIndex = 2
        .Font.Bold = True
    End With


    '##############
    '�e�X�g���{�҂̎擾
    '##############
    Cells(p_int�J�n�s + 2, 1).Value = "�S��"
    
    '�S����/�����҂��Z�b�g
    Keys = p_dic�����o�[.Keys
   
    For j = 0 To p_dic�����o�[.Count - 1
        Cells(p_int�J�n�s + 3 + j, 1).Value = Keys(j)
    Next j
    
    '���v
    Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1).Value = "���v"
  
     '19/03/26 Mod Start
'    '�P�T�ԒP�ʂɁu�\��v�u���сv�񍀖ڂ�p�ӂ���
'    l_intNissu = p_intNissu
'    l_wrkDate = l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Value
'    l_intWeeks = (m_intEndCoumn - m_intStartColumn) / 4 / 7
'
'    For k = 0 To l_intWeeks - 1
'        For m = 0 To 1
'            Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Value = l_wrkDate + 4   '���j�����T����Ƃ���
'
'            If ((2 + m) Mod 2) = 0 Then      '�\��
'                Cells(p_int�J�n�s + 2, 2 + (2 * k) + m).Value = "�\��"
'                '�S����
'                If p_int�J�n�s = 2 Then     '�e�X�g���{
'                    For j = 0 To p_dic�����o�[.Count - 1
'                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 6).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """���{����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """�\��""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
'                    Next j
'                Else                        '�e�X�g����
'                    For j = 0 To p_dic�����o�[.Count - 1
'                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 11).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """��������""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """�\��""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
'                    Next j
'
'                End If
'                '���v
'                Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + (2 * k) + m) = _
'                "=SUM(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count - 1, 2 + (2 * k) + m).Address + ")"
'
'            ElseIf ((2 + m) Mod 2) = 1 Then  '����
'                Cells(p_int�J�n�s + 2, 2 + (2 * k) + m).Value = "����"
'                '�S����
'                If p_int�J�n�s = 2 Then     '�e�X�g���{
'                    For j = 0 To p_dic�����o�[.Count - 1
'                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 6).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """���{����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """����""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
'                    Next j
'                Else                        '�e�X�g����
'                    For j = 0 To p_dic�����o�[.Count - 1
'                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
'                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 11).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """��������""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """����""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
'                    Next j
'                End If
'                '���v
'                Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + (2 * k) + m) = _
'                "=SUM(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count - 1, 2 + (2 * k) + m).Address + ")"
'            End If
'        Next m
'        l_wrkDate = l_wrkDate + 7
'    Next k
    
'    '�u���ѓ��ɂ��Ǘ��v�񍀖ڂ�p�ӂ���
'    With Range(Cells(p_int�J�n�s + 1, 1 + l_intWeeks * 2 + 1), Cells(p_int�J�n�s + 1, 1 + l_intWeeks * 2 + 2))
'        .Merge
'        .Value = "���ѓ��ɂ��Ǘ�" + vbCrLf + "(�w�E�L���Ɋւ�炸" + vbCrLf + "����̎��ѓ��͓��ŊǗ�)"
'        .WrapText = True
'        .ColumnWidth = 15
'        .RowHeight = 40
'    End With
'
'    Cells(p_int�J�n�s + 2, 1 + l_intWeeks * 2 + 1) = "�i����"
'    '�S���� + ���v
'    For j = 0 To p_dic�����o�[.Count
'        With Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2 + 1)
'            .Value = "=SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1 + l_intWeeks * 2).Address + ")*(" + """����""" + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 1 + l_intWeeks * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2).Address + ")" + _
'            "/SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1 + l_intWeeks * 2).Address + ")*(" + """�\��""" + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 1 + l_intWeeks * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2).Address + ")"
'            .NumberFormatLocal = "0%"
'        End With
'    Next j
'
'    Cells(p_int�J�n�s + 2, 1 + l_intWeeks * 2 + 2) = "������"
'    '�S���� + ���v
'    For j = 0 To p_dic�����o�[.Count
'        With Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2 + 2)
'            .Value = "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2).Address + "/" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2 - 1).Address
'            .NumberFormatLocal = "0%"
'        End With
'    Next j
'
'    '�w�i�F
'    Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 1, 1 + l_intWeeks * 2 + 2)).Interior.ColorIndex = 40
'    Range(Cells(p_int�J�n�s + 2, 1), Cells(p_int�J�n�s + 2, 1 + l_intWeeks * 2 + 2)).Interior.ColorIndex = 38
'    '�S���� + ���v
'    For j = 0 To p_dic�����o�[.Count
'        If (j Mod 2) = 1 Then
'            Range(Cells(p_int�J�n�s + 3 + j, 1), Cells(p_int�J�n�s + 3 + j, 1 + l_intWeeks * 2 + 2)).Interior.ColorIndex = 24
'        End If
'    Next j
'
'    '�r���`��
'    Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1 + l_intWeeks * 2 + 2)).Borders.LineStyle = xlContinuous
'
'    '�����`��
'    '�O�g
'    With Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1 + l_intWeeks * 2 + 2))
'        .Borders(xlEdgeRight).Weight = xlMedium
'        .Borders(xlEdgeLeft).Weight = xlMedium
'        .Borders(xlEdgeTop).Weight = xlMedium
'        .Borders(xlEdgeBottom).Weight = xlMedium
'    End With
'    '����
'    Range(Cells(p_int�J�n�s + 2, 1), Cells(p_int�J�n�s + 2, 1 + l_intWeeks * 2 + 2)).Borders(xlEdgeBottom).Weight = xlMedium
'
'    For k = 0 To l_intWeeks * 2
'        If (k Mod 2) = 0 Then
'            With Range(Cells(p_int�J�n�s + 1, 2 + k), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + k))
'                .Borders(xlEdgeLeft).Weight = xlMedium
'            End With
'        End If
'    Next k
'
'
'    Rows(3).Columns.AutoFit                                         '�Z���̗񕝎����ݒ�
'
'    For n = 1 To l_intWeeks * 2 + 1
'        Cells(3, n).ColumnWidth = Cells(3, n).ColumnWidth * 1.2     '�Z���̗񕝎����ݒ� �~�P�D�Q�{
'    Next n


    '�P���P�ʂɁu�\��v�u���сv�񍀖ڂ�p�ӂ���
    l_intNissu = p_intNissu + 1
    l_wrkDate = l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Value
    l_intWeeks = (m_intEndCoumn - m_intStartColumn) / 4 / 7

    For k = 0 To l_intNissu - 1
        For m = 0 To 1
            Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Value = l_wrkDate   '���

            If ((2 + m) Mod 2) = 0 Then     '�\��
                Cells(p_int�J�n�s + 2, 2 + (2 * k) + m).Value = "�\��"
                '�S����
                If p_int�J�n�s = 2 Then     '�e�X�g���{
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 6).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """���{����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """�\��""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
                    Next j
                Else                        '�e�X�g����
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 11).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """��������""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """�\��""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
                    Next j

                End If
                '���v
                Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count - 1, 2 + (2 * k) + m).Address + ")"
                
            ElseIf ((2 + m) Mod 2) = 1 Then  '����
                Cells(p_int�J�n�s + 2, 2 + (2 * k) + m).Value = "����"
                '�S����
                If p_int�J�n�s = 2 Then     '�e�X�g���{
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 6).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """���{����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """����""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
                    Next j
                Else                        '�e�X�g����
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 11).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """��������""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """����""" + "),�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I
                    Next j
                End If
                '���v
                Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count - 1, 2 + (2 * k) + m).Address + ")"
            End If
        Next m
        l_wrkDate = l_wrkDate + 1
    Next k

    '�u���ѓ��ɂ��Ǘ��v�񍀖ڂ�p�ӂ���
    With Range(Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2 + 1), Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2 + 2))
        .Merge
        .Value = "���ѓ��ɂ��Ǘ�" + vbCrLf + "(�w�E�L���Ɋւ�炸" + vbCrLf + "����̎��ѓ��͓��ŊǗ�)"
        .WrapText = True
        .ColumnWidth = 15
        .RowHeight = 40
    End With
    
    Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 1) = "�i����"
    '�S���� + ���v
    For j = 0 To p_dic�����o�[.Count
        With Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 + 1)
            .Value = "=SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2).Address + ")*(" + """����""" + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2).Address + ")" + _
            "/SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2).Address + ")*(" + """�\��""" + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2).Address + ")"
            .NumberFormatLocal = "0%"
        End With
    Next j

    Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2) = "������"
    '�S���� + ���v
    For j = 0 To p_dic�����o�[.Count
        With Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 + 2)
            .Value = "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2).Address + "/" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 - 1).Address
            .NumberFormatLocal = "0%"
        End With
    Next j
    
    '�w�i�F
    'Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 40
    'Range(Cells(p_int�J�n�s + 2, 1), Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 38
    
    For k = 1 To l_intNissu * 2 + 1 Step 2
        With Range(Cells(p_int�J�n�s + 1, 1 + k), Cells(p_int�J�n�s + 2, 2 + k))
        
        If l_bolChgFrg Then
            .Interior.ColorIndex = 35
            l_bolChgFrg = False
        Else
            .Interior.ColorIndex = 40
            l_bolChgFrg = True
        End If
        
        End With
    Next k
    
    Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 2, 1)).Interior.ColorIndex = 38
    
        
    '�S���� + ���v
    For j = 0 To p_dic�����o�[.Count
        If (j Mod 2) = 1 Then
            Range(Cells(p_int�J�n�s + 3 + j, 1), Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 24
        End If
    Next j

    '�r���`��
    Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1 + l_intNissu * 2 + 2)).Borders.LineStyle = xlContinuous
    
    '�����`��
    '�O�g
    With Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1 + l_intNissu * 2 + 2))
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    '����
    Range(Cells(p_int�J�n�s + 2, 1), Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2)).Borders(xlEdgeBottom).Weight = xlMedium
    
    For k = 0 To l_intNissu * 2
        If (k Mod 2) = 0 Then
            With Range(Cells(p_int�J�n�s + 1, 2 + k), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + k))
                .Borders(xlEdgeLeft).Weight = xlMedium
            End With
        End If
    Next k


    Rows(3).Columns.AutoFit                                         '�Z���̗񕝎����ݒ�
    
    For n = 1 To l_intNissu * 2 + 1
        Cells(3, n).ColumnWidth = Cells(3, n).ColumnWidth * 1.2     '�Z���̗񕝎����ݒ� �~�P�D�Q�{
    Next n

    '��̃O���[�v��
    If p_int�J�n�s <> 2 Then    '�����̕\����鎞�̂ݎ��{
        For k = 0 To l_intNissu * 2
            If (k Mod 14) = 0 And k <> 0 Then
                With Range(Cells(p_int�J�n�s + 1, k - 13 + 1), Cells(p_int�J�n�s + 1, k - 1))
                    .Columns.Group                                      '��\����Ԃɂ����������O���[�v��
                    ActiveSheet.Outline.ShowLevels columnlevels:=1      '�����_�ŃO���[�v������Ă�������\���ɂ���
                End With
            End If
        Next k
    End If
    
    
    '19/03/26 Mod End
    
    
End Sub

'19/04/04 Add Start
'################################
'�\���Ǘ�(�@�\�P��)�V�[�g���쐬����
'################################
Private Sub �\���Ǘ��쐬_�@�\�P��(p_intNissu As Integer)

    Worksheets("�\���Ǘ�_�@�\�P��").Activate

    '�V�[�g������
    Application.DisplayAlerts = False   '�m�F���b�Z�[�W�̔�\��
    Range("3:3").UnMerge                '�Z���̌�������
    Cells.ClearFormats                  '�S�Z���̕\���`���N���A
    Cells.Clear                         '�S�Z���̃N���A
    Cells.ColumnWidth = 5               '�S�Z���̕�
    Cells.RowHeight = 20                '�S�Z���̍���
    ActiveSheet.Columns.ClearOutline    '��̃O���[�v����S�č폜

    '�ϐ��錾
    Dim Dic�S���� As Object
    Dim Dic������ As Object
    Dim Dic�@�\ As Object
    Dim Dic�Ώۋ@�\ As Object
    Dim l_int�e�X�g���{start As Integer
    Dim l_int�e�X�g����start As Integer
    Dim l_int�@�\�����o�[�� As Integer
    Dim l_int�@�\_�S���Ґ� As Integer
    Dim l_int�@�\_�����Ґ� As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Keys_�@�\() As Variant
    Dim Keys_�S����() As Variant
    Dim Keys_������() As Variant
    Dim l_str�@�\ As String
    
    Set Dic�Ώۋ@�\ = CreateObject("Scripting.Dictionary")
    
    '�^�C�g��
    With Cells(1, 1)
        .Value = "�\���Ǘ�(�@�\�P��)"
        .Font.Bold = True
    End With


    '�@�\�̎擾
    Set Dic�@�\ = �@�\�擾()
    Keys_�@�\ = Dic�@�\.Keys                    '�@�\�L�[���Z�b�g
    
    '�e�X�g���{�҂̎擾
    Set Dic�S���� = �����o�[�擾_�@�\�P��("�S����")
    Keys_�S���� = Dic�S����.Keys                '�@�\�L�[���Z�b�g
    
    '�e�X�g�����҂̎擾
    Set Dic������ = �����o�[�擾_�@�\�P��("������")
    Keys_������ = Dic������.Keys                '�@�\�L�[���Z�b�g

    
    '�@�\�P�ʂɃe�X�g���{�A�e�X�g�����̗\���Ǘ��\���쐬����B
    For j = 0 To Dic�@�\.Count - 1
    
        '�@�\���̎擾
        l_str�@�\ = Keys_�@�\(j)
            
        '�ŏ��̋@�\�ɂ��e�X�g���{�̊J�n�s�̐ݒ�
        If j = 0 Then
            l_int�e�X�g���{start = 2
        End If
                
                
        '################
        '## �e�X�g���{ ##
        '################

        '�Ώۋ@�\�̒S���҂𒊏o����B�i���̃e�X�g�����ƁA���@�\�̃e�X�g���{�̊J�n�s�̐ݒ菀���ɂ��Ȃ�j
        For k = 0 To Dic�S����.Count - 1
            If l_str�@�\ = Dic�S����.Item(Keys_�S����(k))(1) Then
                If Not Dic�Ώۋ@�\.exists(Dic�S����.Item(Keys_�S����(k))(0)) Then
                    Dic�Ώۋ@�\.Add Dic�S����.Item(Keys_�S����(k))(0), Null
                    l_int�@�\_�S���Ґ� = l_int�@�\_�S���Ґ� + 1
                End If
            End If
        Next k
        
        '�e�X�g���{�҂̗\���Ǘ��\�쐬(�@�\�P��)
        Call �\���Ǘ����ו\�쐬_�@�\�P��("�e�X�g���{", l_int�e�X�g���{start, Dic�Ώۋ@�\, l_str�@�\, p_intNissu)

        Dic�Ώۋ@�\.RemoveAll
        
        '################
        '## �e�X�g���� ##
        '################
        
        '�e�X�g�����̊J�n�s�̐ݒ�
        '�@�\�P�ʂ̃e�X�g���{�̊J�n�s + (���t,�S��)(2�s��) + �S���Ґ� + ���v�s(1�s��) + ��(2�s��)
        l_int�e�X�g����start = l_int�e�X�g���{start + 2 + l_int�@�\_�S���Ґ� + 1 + 2
        
        
        '�Ώۋ@�\�̐����҂𒊏o����B�i���̋@�\�Ɍ����āA�e�X�g�����̊J�n�s�̐ݒ菀���ɂ��Ȃ�j
        For k = 0 To Dic������.Count - 1
            If l_str�@�\ = Dic������.Item(Keys_������(k))(1) Then
                If Not Dic�Ώۋ@�\.exists(Dic������.Item(Keys_������(k))(0)) Then
                    Dic�Ώۋ@�\.Add Dic������.Item(Keys_������(k))(0), Null
                    l_int�@�\_�����Ґ� = l_int�@�\_�����Ґ� + 1
                End If
            End If
        Next k
        
        '�e�X�g�����҂̗\���Ǘ��\�쐬(�@�\�P��)
        Call �\���Ǘ����ו\�쐬_�@�\�P��("�e�X�g����", l_int�e�X�g����start, Dic�Ώۋ@�\, l_str�@�\, p_intNissu)
        
        Dic�Ώۋ@�\.RemoveAll
        
        '���̋@�\�P�ʂ̃e�X�g���{�̊J�n�s�̐ݒ�
        '�@�\�P�ʂ̃e�X�g�����J�n�s + (���t,�S��)(2�s��) + �����҂̐l�� + ���v�s(1�s��) + �󔒍s(4�s��)
        l_int�e�X�g���{start = l_int�e�X�g����start + 2 + l_int�@�\_�����Ґ� + 1 + 4
        
        l_int�@�\_�S���Ґ� = 0
        l_int�@�\_�����Ґ� = 0
        l_str�@�\ = ""
    Next j
    
    '��̃O���[�v��
    For k = 0 To (p_intNissu + 1) * 2
        If (k Mod 14) = 0 And k <> 0 Then
            With Range(Cells(1, k - 13 + 1), Cells(1, k - 1))
                .Columns.Group                                      '��\����Ԃɂ����������O���[�v��
                ActiveSheet.Outline.ShowLevels columnlevels:=1      '�����_�ŃO���[�v������Ă�������\���ɂ���
            End With
        End If
    Next k


    Cells.HorizontalAlignment = xlCenter
    Columns("A").HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter

        
    Application.DisplayAlerts = True    '�m�F���b�Z�[�W�̕\��
    
    Worksheets("�����Ǘ�").Activate

End Sub


'###########################
'�\���Ǘ����ו\�쐬_�@�\�P��
'###########################
Private Sub �\���Ǘ����ו\�쐬_�@�\�P��(p_str�e�X�g���{_���� As String, p_int�J�n�s As Integer, p_dic�����o�[ As Object, p_str�@�\ As String, p_intNissu As Integer)
    
    
    '�ϐ��錾
    Dim l_sht_�����Ǘ� As Worksheet
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

    '�ϐ�������
    Set l_sht_�����Ǘ� = Sheets("�����Ǘ�")
    Set l_sht_Yojitsu = Sheets("�\���Ǘ�_�@�\�P��")
    
    If p_str�e�X�g���{_���� = "�e�X�g���{" Then     '�e�X�g���{
        With Cells(p_int�J�n�s, 1)
            .Value = p_str�@�\
            .Font.Bold = True
        End With
        With Cells(p_int�J�n�s, 2)
            .Value = "�e�X�g���{"
            .Font.Bold = True
        End With
    Else                                            '�e�X�g����
        With Cells(p_int�J�n�s, 1)
            .Value = p_str�@�\
            .Font.Bold = True
        End With
        With Cells(p_int�J�n�s, 2)
            .Value = "�e�X�g����"
            .Font.Bold = True
        End With
    End If

    '�{�����t
    With Cells(p_int�J�n�s + 1, 1)
        .Value = "=TODAY()"
        .Font.ColorIndex = 2
        .Font.Bold = True
    End With


    '##############
    '�e�X�g���{�҂̎擾
    '##############
    Cells(p_int�J�n�s + 2, 1).Value = "�S��"
    
    '�S����/�����҂��Z�b�g
    Keys = p_dic�����o�[.Keys
   
    For j = 0 To p_dic�����o�[.Count - 1
        Cells(p_int�J�n�s + 3 + j, 1).Value = Keys(j)
    Next j
    
        
    '���v
    Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1).Value = "���v"

    '�P���P�ʂɁu�\��v�u���сv�񍀖ڂ�p�ӂ���
    l_intNissu = p_intNissu + 1
    l_wrkDate = l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Value
    l_intWeeks = (m_intEndCoumn - m_intStartColumn) / 4 / 7

    For k = 0 To l_intNissu - 1
        For m = 0 To 1
            Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Value = l_wrkDate   '���

            If ((2 + m) Mod 2) = 0 Then     '�\��
                Cells(p_int�J�n�s + 2, 2 + (2 * k) + m).Value = "�\��"
                '�S����
                If p_str�e�X�g���{_���� = "�e�X�g���{" Then     '�e�X�g���{
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 6).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """���{����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """�\��""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 4).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s, 1).Address + ")," _
                        + "�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������"
                          
                    Next j
                Else                        '�e�X�g����
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 11).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """��������""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """�\��""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 4).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s, 1).Address + ")," _
                        + "�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I"
                    Next j

                End If
                '���v
                Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count - 1, 2 + (2 * k) + m).Address + ")"
                
            ElseIf ((2 + m) Mod 2) = 1 Then  '����
                Cells(p_int�J�n�s + 2, 2 + (2 * k) + m).Value = "����"
                '�S����
                If p_str�e�X�g���{_���� = "�e�X�g���{" Then     '�e�X�g���{
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 6).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 6).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """���{����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 4).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s, 1).Address + ")," _
                        + "�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I"
                    Next j
                Else                        '�e�X�g����
                    For j = 0 To p_dic�����o�[.Count - 1
                        Cells(p_int�J�n�s + 3 + j, 2 + (2 * k) + m) = _
                        "=SUMPRODUCT((�����Ǘ�!" + l_sht_�����Ǘ�.Cells(1, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(1, m_intEndCoumn).Address + "<=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2 + (2 * k) + m).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 11).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 11).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1).Address + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(2, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(2, m_intEndCoumn).Address + "=" + """��������""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(3, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(3, m_intEndCoumn).Address + "=" + """����""" + ")*(�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, 4).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, 4).Address + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s, 1).Address + ")," _
                        + "�����Ǘ�!" + l_sht_�����Ǘ�.Cells(m_lngStartRow, m_intStartColumn).Address + ":" + l_sht_�����Ǘ�.Cells(m_lngEndRow, m_intEndCoumn).Address + ")" '���֐�������I"
                    Next j
                End If
                '���v
                Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + (2 * k) + m) = _
                "=SUM(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3, 2 + (2 * k) + m).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count - 1, 2 + (2 * k) + m).Address + ")"
            End If
        Next m
        l_wrkDate = l_wrkDate + 1
    Next k

    '�u���ѓ��ɂ��Ǘ��v�񍀖ڂ�p�ӂ���
    With Range(Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2 + 1), Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2 + 2))
        .Merge
        .Value = "���ѓ��ɂ��Ǘ�" + vbCrLf + "(�w�E�L���Ɋւ�炸" + vbCrLf + "����̎��ѓ��͓��ŊǗ�)"
        .WrapText = True
        .ColumnWidth = 15
        .RowHeight = 40
    End With
    
    Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 1) = "�i����"
    '�S���� + ���v
    For j = 0 To p_dic�����o�[.Count
        With Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 + 1)
            .Value = "=SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2).Address + ")*(" + """����""" + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2).Address + ")" + _
            "/SUMPRODUCT((" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + "- WEEKDAY(" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1).Address + ") + 6 =" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 1, 1 + l_intNissu * 2).Address + ")*(" + """�\��""" + "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2).Address + ")," + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 2).Address + ":" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2).Address + ")"
            .NumberFormatLocal = "0%"
        End With
    Next j

    Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2) = "������"
    '�S���� + ���v
    For j = 0 To p_dic�����o�[.Count
        With Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 + 2)
            .Value = "=" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2).Address + "/" + l_sht_Yojitsu.Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 - 1).Address
            .NumberFormatLocal = "0%"
        End With
    Next j
    
    '�w�i�F
    'Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 40
    'Range(Cells(p_int�J�n�s + 2, 1), Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 38
    
    For k = 1 To l_intNissu * 2 + 1 Step 2
        With Range(Cells(p_int�J�n�s + 1, 1 + k), Cells(p_int�J�n�s + 2, 2 + k))
        
        If l_bolChgFrg Then
            .Interior.ColorIndex = 35
            l_bolChgFrg = False
        Else
            .Interior.ColorIndex = 40
            l_bolChgFrg = True
        End If
        
        End With
    Next k
    
    Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 2, 1)).Interior.ColorIndex = 38
    
        
    '�S���� + ���v
    For j = 0 To p_dic�����o�[.Count
        If (j Mod 2) = 1 Then
            Range(Cells(p_int�J�n�s + 3 + j, 1), Cells(p_int�J�n�s + 3 + j, 1 + l_intNissu * 2 + 2)).Interior.ColorIndex = 24
        End If
    Next j

    '�r���`��
    Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1 + l_intNissu * 2 + 2)).Borders.LineStyle = xlContinuous
    
    '�����`��
    '�O�g
    With Range(Cells(p_int�J�n�s + 1, 1), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 1 + l_intNissu * 2 + 2))
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    '����
    Range(Cells(p_int�J�n�s + 2, 1), Cells(p_int�J�n�s + 2, 1 + l_intNissu * 2 + 2)).Borders(xlEdgeBottom).Weight = xlMedium
    
    For k = 0 To l_intNissu * 2
        If (k Mod 2) = 0 Then
            With Range(Cells(p_int�J�n�s + 1, 2 + k), Cells(p_int�J�n�s + 3 + p_dic�����o�[.Count, 2 + k))
                .Borders(xlEdgeLeft).Weight = xlMedium
            End With
        End If
    Next k


    Rows(3).Columns.AutoFit                                         '�Z���̗񕝎����ݒ�
    
    For n = 1 To l_intNissu * 2 + 1
        Cells(3, n).ColumnWidth = Cells(3, n).ColumnWidth * 1.2     '�Z���̗񕝎����ݒ� �~�P�D�Q�{
    Next n
    
    '19/03/26 Mod End

End Sub

Private Function �@�\�擾() As Object

    Dim l_DicKinou As Object
    Dim i As Integer
    Dim l_strKinou As String
    Set l_DicKinou = CreateObject("Scripting.Dictionary")
    
    Worksheets("�����Ǘ�").Activate
    
    For i = m_lngStartRow To m_lngEndRow
        '�@�\���擾
        l_strKinou = Cells(i, 4).Value
        
        If Not l_DicKinou.exists(l_strKinou) Then
            l_DicKinou.Add l_strKinou, Null
        End If
        l_strKinou = ""
    Next i

    Set �@�\�擾 = l_DicKinou
    
    Worksheets("�\���Ǘ�_�@�\�P��").Activate
    
End Function

Private Function �����o�[�擾_�@�\�P��(p_strRoll As String) As Object

    Dim l_Dic As Object
    Dim i As Integer
    Dim j As Integer
    Dim l_strName As String
    Dim l_strKinou As String
    Dim Info(1) As String
    Set l_Dic = CreateObject("Scripting.Dictionary")
    Dim Keys() As Variant
    'l_Dic.RemoveAll
    
    Worksheets("�����Ǘ�").Activate
    
    j = 0
    
    For i = m_lngStartRow To m_lngEndRow
        If p_strRoll = "�S����" Then
            '�S���҂��擾
            l_strName = Cells(i, 6).Value
            l_strKinou = Cells(i, 4).Value
        Else
            '�����҂��擾
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

    Set �����o�[�擾_�@�\�P�� = l_Dic

    
    Worksheets("�\���Ǘ�_�@�\�P��").Activate
    
End Function
'19/04/04 Add End


