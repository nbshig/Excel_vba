Attribute VB_Name = "Module2"
Public Sub VisibleNames()
    Dim name As Object
    For Each name In Names
        If name.Visible = False Then
            name.Visible = True
        End If
    Next
    MsgBox "���ׂĂ̖��O�̒�`��\�����܂����B", vbOKOnly
End Sub
