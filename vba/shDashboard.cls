Option Explicit

Private Sub Worksheet_Activate()
    AutoMAJ
End Sub

' Boutons et noms d'agents (cellules liées à elles-mêmes)
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim c As Range
    Set c = Target.Range.Cells(1, 1)
    Select Case c.Address(False, False)
        Case "I1": MettreAJour: Exit Sub
        Case "K1": OuvrirCarte: Exit Sub
        Case "K5": ChoisirPeriode 1: Exit Sub
        Case "L5": ChoisirPeriode 7: Exit Sub
        Case "M5": ChoisirPeriode 30: Exit Sub
        Case "N5": ChoisirPeriode 0: Exit Sub
    End Select
    If c.Row = 11 And c.Column = 11 Then
        ChoisirAgent "Tous"
    ElseIf c.Row > 11 And c.Column = 2 Then
        ChoisirAgent CStr(c.Value)
    End If
End Sub

' Changement d'agent ou de dates
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("C5,H5,J5")) Is Nothing Then ActualiserDashboard
End Sub
