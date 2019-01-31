Attribute VB_Name = "Module2"
Public Sub VisibleNames()
    Dim name As Object
    For Each name In Names
        If name.Visible = False Then
            name.Visible = True
        End If
    Next
    MsgBox "すべての名前の定義を表示しました。", vbOKOnly
End Sub
