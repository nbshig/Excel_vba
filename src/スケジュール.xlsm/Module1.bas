Attribute VB_Name = "Module1"
Option Explicit

Sub Copy_Task(ByVal p_lngRowCount As Long, ByVal p_targetValue As String)

    Dim l_lngLastRowBeforTarget As Long
    Dim l_lngLastRowB As Long

    
    'A��̃^�[�Q�b�g�s��蒼�O��A����t�̓��͍s���擾
    l_lngLastRowBeforTarget = Cells(p_lngRowCount, 1).End(xlUp).Row
    
    'B���ɂ����āA�l�����͂���Ă���ŏI�s���擾
    l_lngLastRowB = Cells(Rows.Count, 2).End(xlUp).Row

    '�O�T�̃^�X�N���R�s�[&�^�[�Q�b�g�s�Ƀy�[�X�g
    Range(Cells(l_lngLastRowBeforTarget, 1), Cells(l_lngLastRowB, 1)).EntireRow.Copy _
        Destination:=Cells(p_lngRowCount, 1)
    
    'A��^�[�Q�b�g�s�ɒl�����
    Cells(p_lngRowCount, 1) = p_targetValue
    
End Sub


Sub Set_FormatConditions(ByVal p_lngNewLastRowB As Long)

    Dim fc As FormatCondition
    
    '�����t���������폜
    Cells.FormatConditions.Delete
        
    
    '=$I1="����" ��=B��1�s��:I��ŏI�s�͈̔͂Ŕ����D�F
    With Range("I1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$I1=""����""")
        fc.Interior.Color = XlRgbColor.rgbLightGray
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With


    '=$G1="����" ��=B��1�s��:I��ŏI�s�͈̔͂Ŕ����D�F
    With Range("G1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$G1=""����""")
        fc.Interior.Color = XlRgbColor.rgbLightGray
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With


    '=$B1="No" ��=B��1�s��:I��ŏI�s�͈̔͂ŒW���A�N�A�}�����F
    With Range("B1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$B1=""No""")
        fc.Interior.Color = XlRgbColor.rgbMediumAquamarine
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With


    '=AND($B1<>"No", $B1<>"",MOD(ROW(),2)=0) ��=B��1�s��:I��ŏI�s�͈̔͂Ń��l���F
    With Range("B1", "I" & p_lngNewLastRowB)
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($B1<>""No"", $B1<>"""",MOD(ROW(),2)=0)")
        fc.Interior.Color = XlRgbColor.rgbLinen
        fc.ModifyAppliesToRange Range(Cells(1, 2), Cells(p_lngNewLastRowB, 9))
    End With
End Sub
