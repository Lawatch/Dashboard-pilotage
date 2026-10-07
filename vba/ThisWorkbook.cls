Option Explicit

' À l'ouverture : modifications faites sur la carte, mise à jour si une source a changé, vues réécrites
Private Sub Workbook_Open()
    On Error Resume Next
    ImporterCarte
    ThisWorkbook.Worksheets(F_DASH).Activate
    Ouverture
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ArreterVeille
End Sub

' Collage dans un onglet jour ou dans ABONNEMENTS : mise à jour au retour sur DASHBOARD / CONTROLES
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name = F_ABO Or Not EstOngletSysteme(Sh.Name) Then
        Application.EnableEvents = False
        MarquerAActualiser
        Application.EnableEvents = True
    End If
End Sub
