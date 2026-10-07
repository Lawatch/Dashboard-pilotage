Option Explicit

Private Sub Worksheet_Activate()
    AutoMAJ
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim c As Range
    Set c = Target.Range.Cells(1, 1)
    If c.Address(False, False) = "I1" Then
        MettreAJour
    ElseIf c.Column = 9 And CStr(Me.Cells(c.Row, 1).Value) = "SANSABO" Then
        AccepterSuggestion c
    End If
End Sub

' N° d'abonnement saisi en colonne J : rattachement immédiat
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim c As Range
    If Target.Cells.CountLarge > 500 Then Exit Sub
    For Each c In Target.Cells
        If c.Column = 10 And CStr(Me.Cells(c.Row, 1).Value) = "SANSABO" And CStr(c.Value) <> "" Then
            MajRattachements
            Exit Sub
        End If
    Next c
End Sub
