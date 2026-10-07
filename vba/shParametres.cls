Option Explicit

' Bouton « Localiser les adresses et recalculer les km »
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If Target.Range.Cells(1, 1).Address(False, False) = "B16" Then RecalculerKm
End Sub

' Tout réglage modifié est pris en compte au retour sur DASHBOARD
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    MarquerAActualiser
    Application.EnableEvents = True
End Sub
