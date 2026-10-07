'==============================================================================
' Mise à jour automatique : onglets jours -> DATA -> abonnements -> km -> CONTROLES -> DASHBOARD
'
'   AutoMAJ      : appelée à l'ouverture et quand on revient sur DASHBOARD / CONTROLES.
'                  Ne fait rien si aucune source n'a changé.
'   MettreAJour  : bouton « Mettre à jour » (refait tout, affiche le bilan)
'
' Onglets jours : n'importe quel nom ; l'onglet est renommé d'après la date lue
' dans ses lignes (JJMMAA). Un jour présent deux fois : la version la plus à
' droite est gardée, l'autre est renommée « ancien » (ou « doublon » si identique).
'==============================================================================
Option Explicit

Private gEnCours As Boolean

'------------------------------------------------------------------------------
' Points d'entrée
'------------------------------------------------------------------------------
Public Sub MettreAJour()
    MiseAJourComplete True
End Sub

Public Sub AutoMAJ()
    If gEnCours Then Exit Sub
    If BesoinMAJ() Then MiseAJourComplete False
End Sub

' À l'ouverture : mise à jour si une source a changé, sinon simple réécriture des vues
Public Sub Ouverture()
    If BesoinMAJ() Then
        MiseAJourComplete False
    Else
        Application.ScreenUpdating = False
        EcrireControles ""
        ActualiserDashboard
        Application.ScreenUpdating = True
    End If
End Sub

' Une source a-t-elle changé depuis la dernière mise à jour ?
Public Function BesoinMAJ() As Boolean
    Dim ws As Worksheet, nb As Long
    On Error GoTo Oui
    If LireEtat("a_actualiser") = "1" Then BesoinMAJ = True: Exit Function
    If LireEtat("sig_abonnements") <> SignatureAbonnements() Then BesoinMAJ = True: Exit Function
    For Each ws In ThisWorkbook.Worksheets
        If EstOngletSource(ws) Then
            nb = nb + 1
            If LireSignature(ws.Name) <> Signature(ws) Then BesoinMAJ = True: Exit Function
        End If
    Next ws
    If nb <> CLng(Val(LireEtat("nb_onglets"))) Then BesoinMAJ = True
    Exit Function
Oui:
    BesoinMAJ = True
End Function

' Appelée par les onglets sources et PARAMETRES quand on y modifie quelque chose
Public Sub MarquerAActualiser()
    On Error Resume Next
    EcrireEtat "a_actualiser", "1"
End Sub

Public Sub MiseAJourComplete(ByVal manuel As Boolean)
    Dim calcAvant As Long, etape As String, msg As String, msgO As String, msgJ As String
    Dim msgA As String, msgK As String, joursModifies As String, wsRetour As Worksheet, nbJ As Long, dMax As Long
    Dim journal As String, t0 As Single

    If gEnCours Then Exit Sub
    gEnCours = True
    Set wsRetour = ActiveSheet
    calcAvant = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Echec

    t0 = Timer
    etape = "corrections saisies dans CONTROLES"
    Application.StatusBar = "Mise à jour : corrections..."
    RecupererCorrections
    etape = "nommage des onglets jours"
    Application.StatusBar = "Mise à jour : onglets jours..."
    msgO = RenommerOngletsJours()
    Chrono journal, t0, "onglets"
    etape = "lecture des onglets jours"
    msgJ = IntegrerJours(joursModifies, nbJ)
    EcrireEtat "nb_jours", CStr(nbJ)
    Chrono journal, t0, "jours"
    etape = "agents"
    MajAgents
    Chrono journal, t0, "agents"
    etape = "abonnements"
    Application.StatusBar = "Mise à jour : abonnements..."
    msgA = TraiterAbonnements()
    Chrono journal, t0, "abonnements"
    etape = "kilomètres"
    If joursModifies <> "|" Or LireEtat("sig_km") <> SignatureKm() Then
        Application.StatusBar = "Mise à jour : kilomètres..."
        msgK = CalculerKm(IIf(LireEtat("sig_km") <> SignatureKm(), "*", joursModifies))
        EcrireEtat "sig_km", SignatureKm()
    End If
    Chrono journal, t0, "km"
    etape = "contrôles"
    Application.StatusBar = "Mise à jour : contrôles..."
    EcrireControles msgO & msgJ
    Chrono journal, t0, "contrôles"
    etape = "tableau de bord"
    EcrireEtat "maj", Format(Now, "dd/mm/yyyy hh:mm")
    dMax = DernierJour()
    If CStr(dMax) <> LireEtat("dmax") Or Not EstUneDate(ThisWorkbook.Worksheets(F_DASH).Range(CELL_DU).Value) Then PeriodeParDefaut
    EcrireEtat "dmax", CStr(dMax)
    ActualiserDashboard
    Chrono journal, t0, "dashboard"
    MemoriserSignatures
    EcrireEtat "a_actualiser", "0"
    EcrireEtat "chrono", journal

Fin:
    On Error Resume Next
    Application.Calculation = calcAvant
    Application.StatusBar = False
    Application.EnableEvents = True
    wsRetour.Activate
    Application.ScreenUpdating = True
    gEnCours = False
    On Error GoTo 0
    If manuel Then
        msg = "Mise à jour terminée."
        If msgO <> "" Then msg = msg & vbCrLf & vbCrLf & msgO
        If msgJ <> "" Then msg = msg & vbCrLf & vbCrLf & msgJ
        If msgA <> "" Then msg = msg & vbCrLf & vbCrLf & msgA
        If msgK <> "" Then msg = msg & vbCrLf & vbCrLf & msgK
        MsgBox msg, IIf(InStr(1, msg, "erreur", vbTextCompare) > 0, vbExclamation, vbInformation), "Mise à jour"
    ElseIf InStr(1, msgA & msgK, "erreur", vbTextCompare) > 0 Then
        Application.StatusBar = "Mise à jour : " & Replace(msgA & " " & msgK, vbCrLf, " ")
    End If
    Exit Sub

Echec:
    msgA = msgA & IIf(msgA <> "", vbCrLf, "") & "Erreur pendant l'étape « " & etape & " » : " & Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets(F_DASH).Range("B3").Value = ChrW(9888) & " " & msgA
    Resume Fin
End Sub

' Durée de chaque étape (LISTES, « chrono ») : utile si la mise à jour devient lente
Private Sub Chrono(ByRef s As String, ByRef t0 As Single, ByVal etape As String)
    s = s & etape & " " & Format(Timer - t0, "0.0") & " s | "
    t0 = Timer
End Sub

'------------------------------------------------------------------------------
' Onglets sources
'------------------------------------------------------------------------------
' Onglet contenant un extract de validations (en-tête « Nom Collaborateur » en colonne G)
Public Function EstOngletSource(ws As Worksheet) As Boolean
    If EstOngletSysteme(ws.Name) Or EstOngletEcarte(ws.Name) Then Exit Function
    EstOngletSource = (LigneEntete(ws) > 0)
End Function

' Ligne d'en-tête (0 si l'onglet n'est pas un extract) ; format = 1 extract brut, 2 format DATA
Public Function LigneEntete(ws As Worksheet, Optional ByRef fmtOut As Long) As Long
    Dim i As Long, v As Variant
    fmtOut = 0
    v = ws.Range("A1:G10").Value
    For i = 1 To 10
        If Trim$(txt(v(i, 7))) = "Nom Collaborateur" Then
            LigneEntete = i: fmtOut = 1: Exit Function
        ElseIf Trim$(txt(v(i, 3))) = "Agent" And Trim$(txt(v(i, 1))) = "Date" Then
            LigneEntete = i: fmtOut = 2: Exit Function
        End If
    Next i
End Function

' Empreinte : taille + total des durées prévues + nombre de pointages + agents
Public Function Signature(ws As Worksheet) As String
    Dim s As String
    On Error Resume Next
    s = DerniereLigne(ws) & "x" & DerniereColonne(ws)
    s = s & "|" & Format(Application.WorksheetFunction.Sum(ws.Range("E:E")), "0.00")
    s = s & "|" & Application.WorksheetFunction.CountA(ws.Range("H:H"))
    s = s & "|" & Application.WorksheetFunction.CountA(ws.Range("G:G"))
    s = s & "|" & Format(Application.WorksheetFunction.Sum(ws.Range("F:F")), "0")
    Signature = s
End Function

Public Function SignatureAbonnements() As String
    Dim ws As Worksheet, s As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(F_ABO)
    s = DerniereLigne(ws) & "x" & DerniereColonne(ws)
    s = s & "|" & Format(Application.WorksheetFunction.Sum(ws.Range("B:B")), "0")
    s = s & "|" & Application.WorksheetFunction.CountA(ws.Range("A:Z"))
    SignatureAbonnements = s
End Function

' Adresses et réglages qui changent les km
Public Function SignatureKm() As String
    Dim ws As Worksheet, s As String, lastR As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(F_PARAM)
    lastR = DerniereLigne(ws)
    s = txt(ws.Range(P_ADR_AGENCE).Value) & "|" & txt(ws.Range(P_RETOUR).Value) & "|" & txt(ws.Range(P_RAYON).Value)
    If lastR >= P_AG1 Then
        s = s & "|" & Application.WorksheetFunction.CountA(ws.Range("D" & P_AG1 & ":D" & lastR))
        s = s & "|" & Format(Application.WorksheetFunction.Sum(ws.Range("H" & P_AG1 & ":I" & lastR)), "0.00000")
    End If
    SignatureKm = s
End Function

' Dates (numéros de série) présentes dans un onglet source : "|46286|46287|"
Private Function DatesOnglet(ws As Worksheet, ByRef dMin As Long, ByRef dMax As Long) As String
    Dim h As Long, fmt As Long, lastR As Long, v As Variant, i As Long, d As Long, s As String, colAg As Long
    s = "|": dMin = 0: dMax = 0
    h = LigneEntete(ws, fmt)
    If h = 0 Then DatesOnglet = s: Exit Function
    colAg = IIf(fmt = 1, 7, 3)
    lastR = DerniereLigne(ws)
    If lastR <= h Then DatesOnglet = s: Exit Function
    v = ws.Range(ws.Cells(h + 1, 1), ws.Cells(lastR, colAg)).Value
    For i = 1 To UBound(v, 1)
        If Trim$(txt(v(i, colAg))) <> "" Then
            d = SerialDate(v(i, 1))
            If d > 0 Then
                If InStr(s, "|" & d & "|") = 0 Then s = s & d & "|"
                If dMin = 0 Or d < dMin Then dMin = d
                If d > dMax Then dMax = d
            End If
        End If
    Next i
    DatesOnglet = s
End Function

' Renomme chaque onglet source d'après ses dates. Renvoie le bilan ("" si rien à signaler)
Public Function RenommerOngletsJours() As String
    Dim ws As Worksheet, autre As Worksheet, cible As String, dMin As Long, dMax As Long
    Dim bilan As String, ancienNom As String, sigWs As String, sigAutre As String
    For Each ws In ThisWorkbook.Worksheets
        If EstOngletSource(ws) Then
            DatesOnglet ws, dMin, dMax
            If dMin = 0 Then
                ColorerOnglet ws, RGB(198, 40, 40)
            Else
                If dMin = dMax Then cible = JJMMAA(dMin) Else cible = JJMMAA(dMin) & "-" & JJMMAA(dMax)
                If StrComp(ws.Name, cible, vbTextCompare) <> 0 Then
                    ancienNom = ws.Name
                    If Not ExisteOnglet(cible) Then
                        ws.Name = cible
                        bilan = bilan & vbCrLf & "Onglet « " & ancienNom & " » renommé « " & cible & " »."
                    Else
                        Set autre = ThisWorkbook.Worksheets(cible)
                        sigWs = Signature(ws): sigAutre = Signature(autre)
                        If sigWs = sigAutre Then
                            ws.Name = NomLibre(cible & " doublon")
                            ColorerOnglet ws, RGB(166, 166, 166)
                            bilan = bilan & vbCrLf & "Onglet « " & ancienNom & " » : copie identique du " & _
                                    Format(dMin, "dd/mm") & ", mis de côté (« " & ws.Name & " »), vous pouvez le supprimer."
                        ElseIf (LireSignature(cible) = sigAutre And LireSignature(ancienNom) <> sigWs) Or ws.Index > autre.Index Then
                            ' nouvelle version du même jour : elle remplace l'ancienne
                            autre.Name = NomLibre(cible & " ancien")
                            ColorerOnglet autre, RGB(166, 166, 166)
                            ws.Name = cible
                            bilan = bilan & vbCrLf & "Le " & Format(dMin, "dd/mm") & " a été recollé (onglet « " & ancienNom & _
                                    " ») : la nouvelle version remplace l'ancienne, renommée « " & autre.Name & " »."
                        Else
                            ws.Name = NomLibre(cible & " doublon")
                            ColorerOnglet ws, RGB(166, 166, 166)
                            bilan = bilan & vbCrLf & "Onglet « " & ancienNom & " » : le " & Format(dMin, "dd/mm") & _
                                    " existe déjà (« " & cible & " »), cet onglet est mis de côté (« " & ws.Name & " »)."
                        End If
                    End If
                End If
            End If
        End If
    Next ws
    If bilan <> "" Then RenommerOngletsJours = "Onglets jours :" & bilan
    OrdonnerOngletsJours
End Function

' Onglets jours rangés du plus récent au plus ancien, après les onglets principaux
Private Sub OrdonnerOngletsJours()
    Dim ws As Worksheet, noms() As String, ix() As Long, n As Long, i As Long, apres As Worksheet, d1 As Long, d2 As Long
    On Error GoTo Fin
    ReDim noms(1 To ThisWorkbook.Worksheets.Count): ReDim ix(1 To ThisWorkbook.Worksheets.Count)
    For Each ws In ThisWorkbook.Worksheets
        If EstOngletSource(ws) Then
            DatesOnglet ws, d1, d2
            n = n + 1
            noms(n) = Format(99999 - d2, "00000") & "|" & ws.Name
            ix(n) = n
        End If
    Next ws
    If n = 0 Then Exit Sub
    TriTexte noms, ix, 1, n
    Set apres = ThisWorkbook.Worksheets(F_ABO)
    For i = 1 To n
        Set ws = ThisWorkbook.Worksheets(Mid$(noms(i), InStr(noms(i), "|") + 1))
        If ws.Index <> apres.Index + 1 Then ws.Move After:=apres
        Set apres = ws
    Next i
    ' onglets mis de côté (doublon, ancienne version) : à la suite
    For Each ws In ThisWorkbook.Worksheets
        If EstOngletEcarte(ws.Name) Then
            If ws.Index <> apres.Index + 1 Then ws.Move After:=apres
            Set apres = ws
        End If
    Next ws
Fin:
End Sub

'------------------------------------------------------------------------------
' DATA : reconstruit à partir de tous les onglets sources
'------------------------------------------------------------------------------
Private Function IntegrerJours(ByRef joursModifies As String, ByRef nbJours As Long) As String
    Dim ws As Worksheet, wsData As Worksheet, blocs As New Collection, noms() As String, nbs() As Long
    Dim nS As Long, i As Long, j As Long, k As Long, n As Long, total As Long, bloc As Variant
    Dim proprio As New Collection, d As Long, p As Long, sortie() As Variant, lastR As Long
    Dim alertes As String, conflits As New Collection, modifie As Boolean, msgErr As String, nb As Long, nbOk As Long

    Set wsData = ThisWorkbook.Worksheets(F_DATA)
    joursModifies = "|"
    ReDim noms(1 To ThisWorkbook.Worksheets.Count): ReDim nbs(1 To ThisWorkbook.Worksheets.Count)

    ' 1) Lecture des onglets
    For Each ws In ThisWorkbook.Worksheets
        If EstOngletSource(ws) Then
            msgErr = LireOnglet(ws, bloc, nb, nbOk)
            If msgErr <> "" Then
                alertes = alertes & vbCrLf & "Onglet « " & ws.Name & " » non lu : " & msgErr
                ColorerOnglet ws, RGB(198, 40, 40)
            Else
                nS = nS + 1
                noms(nS) = ws.Name: nbs(nS) = nb
                blocs.Add bloc
                ColorerOnglet ws, RGB(46, 125, 50)
                modifie = (LireSignature(ws.Name) <> Signature(ws))
                ' 2) Chaque jour appartient à un seul onglet (priorité à l'onglet qui porte son nom)
                For i = 1 To nb
                    d = bloc(i, D_DATE)
                    p = Idx(proprio, CStr(d))
                    If p = 0 Then
                        Ajout proprio, CStr(d), nS
                    ElseIf p <> nS Then
                        If ws.Name = JJMMAA(d) And noms(p) <> JJMMAA(d) Then
                            proprio.Remove "k" & d: Ajout proprio, CStr(d), nS
                            Ajout conflits, CStr(d), noms(p)
                        Else
                            Ajout conflits, CStr(d), ws.Name
                        End If
                    End If
                    If modifie Then
                        If InStr(joursModifies, "|" & d & "|") = 0 Then joursModifies = joursModifies & d & "|"
                    End If
                Next i
                bloc = Empty            ' le prochain onglet est lu dans un nouveau tableau
            End If
        End If
    Next ws
    ' 3) Assemblage
    For k = 1 To nS
        bloc = blocs(k)
        For i = 1 To nbs(k)
            If Idx(proprio, CStr(bloc(i, D_DATE))) = k Then total = total + 1
        Next i
    Next k
    lastR = DerniereLigne(wsData)
    If lastR >= 2 Then wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastR, NB_COLS)).ClearContents
    nbJours = proprio.Count
    If total > 0 Then
        ReDim sortie(1 To total, 1 To NB_COLS)
        n = 0
        For k = 1 To nS
            bloc = blocs(k)
            For i = 1 To nbs(k)
                If Idx(proprio, CStr(bloc(i, D_DATE))) = k Then
                    n = n + 1
                    For j = 1 To NB_COLS
                        sortie(n, j) = bloc(i, j)
                    Next j
                End If
            Next i
        Next k
        With wsData.Range(wsData.Cells(2, 1), wsData.Cells(total + 1, NB_COLS))
            .Columns(D_REF).NumberFormat = "@"
            .Columns(D_ONGLET).NumberFormat = "@"
            .Value = sortie
            .Sort Key1:=.Cells(1, D_DATE), Order1:=xlAscending, Key2:=.Cells(1, D_AGENT), Order2:=xlAscending, _
                  Key3:=.Cells(1, D_DEBR), Order3:=xlAscending, Header:=xlNo, MatchCase:=False
        End With
    End If
    FormaterData wsData

    ' jours retirés (onglet supprimé) : à retirer aussi des tournées
    If conflits.Count > 0 Then
        alertes = alertes & vbCrLf & conflits.Count & " jour(s) présent(s) dans plusieurs onglets : un seul onglet est pris en compte (voir CONTROLES)."
    End If
    EcrireConflitsJours conflits
    If alertes <> "" Then IntegrerJours = "Onglets jours :" & alertes
End Function

' Mémorise les doublons de jours pour CONTROLES (LISTES, colonnes L:M)
Private Sub EcrireConflitsJours(conflits As Collection)
    Dim wsL As Worksheet, i As Long, lastR As Long
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    lastR = DerniereLigne(wsL)
    wsL.Range("L2:M" & Application.WorksheetFunction.Max(2, lastR)).ClearContents
    wsL.Range("L1:M1").Value = Array("Jour en double", "Onglet ignoré")
    ' la collection ne donne pas ses clés : on la reparcourt via DATA
    Dim v As Variant, d As Long, vus As New Collection, r As Long
    lastR = DerniereLigne(ThisWorkbook.Worksheets(F_DATA))
    If lastR < 2 Or conflits.Count = 0 Then Exit Sub
    v = ThisWorkbook.Worksheets(F_DATA).Range("A2:A" & lastR).Value
    r = 1
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, 1))
        If d > 0 Then
            If Existe(conflits, CStr(d)) And AjoutUnique(vus, CStr(d)) Then
                r = r + 1
                wsL.Cells(r, 12).Value = CDate(d)
                wsL.Cells(r, 13).Value = Valeur(conflits, CStr(d))
            End If
        End If
    Next i
End Sub

' Lit un onglet source. Renvoie "" si OK, sinon le message d'erreur.
Private Function LireOnglet(ws As Worksheet, ByRef sortie As Variant, ByRef nb As Long, ByRef nbOk As Long) As String
    Dim lastR As Long, lastC As Long, v As Variant, c(1 To 16) As Long, fmt As Long
    Dim i As Long, entete As Long, agent As String, d As Long, typ As String
    Dim debR As Variant, finR As Variant, dureeR As Variant, dureeP As Variant, chantier As String
    Const K_DATE = 1, K_DEBP = 2, K_FINP = 3, K_DUREE = 4, K_REF = 5, K_AGENT = 6, K_DEBR = 7, K_FINR = 8
    Const K_DUREER = 9, K_PROJET = 10, K_CHANTIER = 11, K_LIBELLE = 12, K_LAT = 13, K_LON = 14, K_CAT = 15, K_SOUSCAT = 16

    nb = 0: nbOk = 0
    entete = LigneEntete(ws, fmt)
    If entete = 0 Then LireOnglet = "format non reconnu (la colonne G doit contenir « Nom Collaborateur »)": Exit Function
    lastR = DerniereLigne(ws)
    If lastR <= entete Then LireOnglet = "onglet vide": Exit Function
    lastC = DerniereColonne(ws)
    If lastC < 32 Then lastC = 32
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    If fmt = 1 Then
        c(K_DATE) = 1: c(K_DEBP) = 2: c(K_FINP) = 3: c(K_DUREE) = 5: c(K_REF) = 6: c(K_AGENT) = 7
        c(K_DEBR) = 8: c(K_FINR) = 9: c(K_DUREER) = 12: c(K_PROJET) = 15: c(K_CHANTIER) = 16
        c(K_LIBELLE) = 17: c(K_LAT) = 20: c(K_LON) = 21: c(K_CAT) = 30: c(K_SOUSCAT) = 31
    Else
        c(K_DATE) = 1: c(K_REF) = 2: c(K_AGENT) = 3: c(K_PROJET) = 4: c(K_CHANTIER) = 5: c(K_LIBELLE) = 6
        c(K_CAT) = 7: c(K_SOUSCAT) = 8: c(K_DEBP) = 9: c(K_FINP) = 10: c(K_DUREE) = 11: c(K_DEBR) = 12
        c(K_FINR) = 13: c(K_DUREER) = 14: c(K_LAT) = 16: c(K_LON) = 17
    End If

    ReDim sortie(1 To lastR, 1 To NB_COLS)
    For i = entete + 1 To lastR
        agent = Trim$(txt(v(i, c(K_AGENT))))
        d = SerialDate(v(i, c(K_DATE)))
        ' on ignore les lignes sans agent, la 2e ligne d'en-tête et la ligne de total « Nombre »
        If agent <> "" And d > 0 Then
            nb = nb + 1
            debR = Heure(v(i, c(K_DEBR)))
            finR = Heure(v(i, c(K_FINR)))
            dureeP = Nombre(v(i, c(K_DUREE)))
            If Not IsEmpty(dureeP) Then dureeP = Round(dureeP, 2) Else dureeP = 0
            typ = UCase$(Trim$(txt(v(i, c(K_SOUSCAT)))))
            dureeR = Empty
            If Not IsEmpty(debR) And Not IsEmpty(finR) Then
                nbOk = nbOk + 1
                dureeR = Nombre(v(i, c(K_DUREER)))
                If IsEmpty(dureeR) Then
                    dureeR = (finR - debR) * 24
                ElseIf dureeR < 0 Then
                    dureeR = (finR - debR) * 24
                End If
                If dureeR < 0 Then dureeR = dureeR + 24
                dureeR = Round(dureeR, 2)
                sortie(nb, D_POINTEE) = 1
                If typ = CONTENEURS Then
                    sortie(nb, D_HRET) = dureeP            ' sortie le matin / rentrée le soir : durée non mesurable
                    sortie(nb, D_STATUT) = "OK"
                Else
                    sortie(nb, D_HRET) = dureeR
                    If dureeP >= 0.25 And dureeR < dureeP * 0.25 Then
                        sortie(nb, D_STATUT) = "Trop courte"
                    ElseIf dureeR > dureeP * 3 And dureeR - dureeP > 1 Then
                        sortie(nb, D_STATUT) = "Trop longue"
                    Else
                        sortie(nb, D_STATUT) = "OK"
                    End If
                End If
            Else
                sortie(nb, D_POINTEE) = 0
                sortie(nb, D_HRET) = 0
                If Not IsEmpty(debR) Then
                    sortie(nb, D_STATUT) = "Début seul"
                ElseIf Not IsEmpty(finR) Then
                    sortie(nb, D_STATUT) = "Fin seule"
                Else
                    sortie(nb, D_STATUT) = "Non pointée"
                End If
            End If
            chantier = Trim$(txt(v(i, c(K_CHANTIER))))
            If chantier = "" Then chantier = Trim$(txt(v(i, c(K_LIBELLE))))

            sortie(nb, D_DATE) = d
            sortie(nb, D_REF) = Trim$(txt(v(i, c(K_REF))))
            sortie(nb, D_AGENT) = agent
            sortie(nb, D_PROJET) = UCase$(Trim$(txt(v(i, c(K_PROJET)))))
            sortie(nb, D_CHANTIER) = chantier
            sortie(nb, D_LIBELLE) = Trim$(txt(v(i, c(K_LIBELLE))))
            sortie(nb, D_CAT) = Trim$(txt(v(i, c(K_CAT))))
            sortie(nb, D_TYPE) = typ
            sortie(nb, D_DEBP) = Heure(v(i, c(K_DEBP)))
            sortie(nb, D_FINP) = Heure(v(i, c(K_FINP)))
            sortie(nb, D_DUREEP) = dureeP
            sortie(nb, D_DEBR) = debR
            sortie(nb, D_FINR) = finR
            sortie(nb, D_DUREER) = dureeR
            sortie(nb, D_LAT) = Nombre(v(i, c(K_LAT)))
            sortie(nb, D_LON) = Nombre(v(i, c(K_LON)))
            sortie(nb, D_ONGLET) = ws.Name
            sortie(nb, D_COMPTE) = 1
        End If
    Next i
    If nb = 0 Then LireOnglet = "aucune prestation trouvée": Exit Function
    MarquerClicsSimultanes sortie, nb
End Function

' Plusieurs prestations (hors conteneurs) d'un même agent pointées début ET fin à la même minute :
' impossible sur place, les clics ont été faits d'un coup
Private Sub MarquerClicsSimultanes(sortie As Variant, ByVal nb As Long)
    Dim i As Long, cle As String, cpt As New Collection, x As Long
    For i = 1 To nb
        If sortie(i, D_POINTEE) = 1 And sortie(i, D_TYPE) <> CONTENEURS Then
            cle = sortie(i, D_DATE) & "|" & sortie(i, D_AGENT) & "|" & Minutes(sortie(i, D_DEBR)) & "|" & Minutes(sortie(i, D_FINR))
            x = Idx(cpt, cle)
            If x > 0 Then cpt.Remove "k" & cle
            cpt.Add x + 1, "k" & cle
        End If
    Next i
    For i = 1 To nb
        If sortie(i, D_POINTEE) = 1 And sortie(i, D_TYPE) <> CONTENEURS Then
            cle = sortie(i, D_DATE) & "|" & sortie(i, D_AGENT) & "|" & Minutes(sortie(i, D_DEBR)) & "|" & Minutes(sortie(i, D_FINR))
            If Idx(cpt, cle) > 1 Then sortie(i, D_STATUT) = "Clics simultanés"
        End If
    Next i
End Sub

Private Sub FormaterData(wsData As Worksheet)
    With wsData
        .Range("A1").Resize(1, NB_COLS).Value = Array("Date", "Réf agent", "Agent", "Projet", "Chantier", "Libellé RDV", _
            "Catégorie", "Sous-catégorie", "Début prévu", "Fin prévue", "Durée prévue (h)", "Début clic", "Fin clic", _
            "Durée pointée (h)", "Pointée (1/0)", "Latitude", "Longitude", "Onglet", "Compté (1/0)", "N° abonnement", _
            "Rattachement", "Clé chantier", "Heures retenues", "Statut pointage")
        .Columns("A").NumberFormat = "dd/mm/yyyy"
        .Columns("I:J").NumberFormat = "hh:mm"
        .Columns("L:M").NumberFormat = "hh:mm"
        .Columns("K").NumberFormat = "0.00"
        .Columns("N").NumberFormat = "0.00"
        .Columns("W").NumberFormat = "0.00"
        .Range("A1").NumberFormat = "General"
        .Range("I1:N1").NumberFormat = "General"
        .Range("A1").Resize(1, NB_COLS).Font.Bold = True
    End With
End Sub

'------------------------------------------------------------------------------
' Agents : liste déroulante, tableau de PARAMETRES, colonne « Compté » de DATA
'------------------------------------------------------------------------------
Public Sub MajAgents()
    Dim wsD As Worksheet, wsL As Worksheet, wsP As Worksheet, lastR As Long, v As Variant, i As Long, j As Long
    Dim agents As New Collection, noms() As String, ix() As Long, n As Long, arr As Variant
    Dim vus As New Collection, sortie() As Variant, nP As Long, nom As String, pv As Variant, exclus As New Collection
    Set wsD = ThisWorkbook.Worksheets(F_DATA)
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)

    ' agents présents dans les données, triés
    lastR = DerniereLigne(wsD)
    If lastR >= 2 Then
        v = wsD.Range(wsD.Cells(2, D_AGENT), wsD.Cells(lastR, D_AGENT)).Value
        ReDim noms(1 To UBound(v, 1)): ReDim ix(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            nom = Trim$(txt(v(i, 1)))
            If nom <> "" Then
                If AjoutUnique(agents, nom) Then n = n + 1: noms(n) = nom: ix(n) = n
            End If
        Next i
        If n > 0 Then TriTexte noms, ix, 1, n
    End If
    wsL.Range("A2:A2000").ClearContents
    wsL.Range("A1").Value = "Agents"
    wsL.Range("A2").Value = "Tous"
    If n > 0 Then
        ReDim arr(1 To n, 1 To 1)
        For i = 1 To n: arr(i, 1) = noms(i): Next i
        wsL.Range("A3").Resize(n, 1).NumberFormat = "@"
        wsL.Range("A3").Resize(n, 1).Value = arr
    End If
    On Error Resume Next
    ThisWorkbook.Names("ListeAgents").RefersTo = "='" & F_LISTES & "'!$A$2:$A$" & (n + 2)
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="ListeAgents", RefersTo:="='" & F_LISTES & "'!$A$2:$A$" & (n + 2)
    End If
    On Error GoTo 0

    ' PARAMETRES : on garde les saisies, on ajoute les nouveaux agents
    lastR = DerniereLigne(wsP)
    ReDim sortie(1 To 2000, 1 To 9)
    If lastR >= P_AG1 Then
        pv = wsP.Range(wsP.Cells(P_AG1, 2), wsP.Cells(lastR, 10)).Value
        For i = 1 To UBound(pv, 1)
            nom = Trim$(txt(pv(i, 1)))
            If nom <> "" And Not Existe(vus, nom) And nP < 2000 Then
                nP = nP + 1: Ajout vus, nom, nP
                For j = 1 To 9: sortie(nP, j) = pv(i, j): Next j
                sortie(nP, 1) = nom
            End If
        Next i
    End If
    For i = 1 To n
        If Not Existe(vus, noms(i)) And nP < 2000 Then
            nP = nP + 1: Ajout vus, noms(i), nP
            sortie(nP, 1) = noms(i)
        End If
    Next i
    For i = 1 To nP
        If Trim$(txt(sortie(i, 2))) = "" Then sortie(i, 2) = IIf(EstAgentFictif(CStr(sortie(i, 1))), "Non", "Oui")
        For j = 4 To 6
            If Trim$(txt(sortie(i, j))) = "" Then sortie(i, j) = "Non"
        Next j
    Next i
    If lastR >= P_AG1 Then wsP.Range(wsP.Cells(P_AG1, 2), wsP.Cells(lastR, 10)).ClearContents
    If nP > 0 Then
        With wsP.Range(wsP.Cells(P_AG1, 2), wsP.Cells(P_AG1 + nP - 1, 10))
            .Value = sortie
            .Sort Key1:=.Cells(1, 1), Order1:=xlAscending, Header:=xlNo, MatchCase:=False
        End With
        For i = 1 To nP
            If Not OuiNon(sortie(i, 2), True) Then Ajout exclus, CStr(sortie(i, 1)), 1
        Next i
    End If

    ' DATA!S : agent compté dans les indicateurs (1/0)
    lastR = DerniereLigne(wsD)
    If lastR >= 2 Then
        v = wsD.Range(wsD.Cells(2, D_AGENT), wsD.Cells(lastR, D_AGENT)).Value
        ReDim arr(1 To UBound(v, 1), 1 To 1)
        For i = 1 To UBound(v, 1)
            arr(i, 1) = IIf(Existe(exclus, Trim$(txt(v(i, 1)))), 0, 1)
        Next i
        wsD.Range(wsD.Cells(2, D_COMPTE), wsD.Cells(lastR, D_COMPTE)).Value = arr
    End If
End Sub

'------------------------------------------------------------------------------
' État (LISTES, colonnes I:J) et signatures des onglets (LISTES, colonnes C:G)
'------------------------------------------------------------------------------
Public Function LireEtat(ByVal cle As String) As String
    Dim wsL As Worksheet, i As Long
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    For i = 2 To 60
        If CStr(wsL.Cells(i, 9).Value) = cle Then LireEtat = CStr(wsL.Cells(i, 10).Value): Exit Function
        If CStr(wsL.Cells(i, 9).Value) = "" Then Exit Function
    Next i
End Function

Public Sub EcrireEtat(ByVal cle As String, ByVal texte As String)
    Dim wsL As Worksheet, i As Long
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    For i = 2 To 60
        If CStr(wsL.Cells(i, 9).Value) = cle Or CStr(wsL.Cells(i, 9).Value) = "" Then
            wsL.Cells(i, 9).Value = cle
            wsL.Cells(i, 10).NumberFormat = "@"
            wsL.Cells(i, 10).Value = texte
            Exit Sub
        End If
    Next i
End Sub

Public Function DernierJour() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(F_DATA)
    If DerniereLigne(ws) >= 2 Then DernierJour = CLng(Application.WorksheetFunction.Max(ws.Range("A:A")))
End Function

Public Function PremierJour() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(F_DATA)
    If DerniereLigne(ws) >= 2 Then PremierJour = CLng(Application.WorksheetFunction.Min(ws.Range("A:A")))
End Function

Private Function LireSignature(ByVal nom As String) As String
    Dim wsL As Worksheet, i As Long
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    i = 2
    Do While CStr(wsL.Cells(i, 3).Value) <> ""
        If CStr(wsL.Cells(i, 3).Value) = nom Then LireSignature = CStr(wsL.Cells(i, 4).Value): Exit Function
        i = i + 1
    Loop
End Function

Private Sub MemoriserSignatures()
    Dim wsL As Worksheet, ws As Worksheet, r As Long, d1 As Long, d2 As Long, nb As Long
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    wsL.Range("C2:G2000").ClearContents
    wsL.Range("C1:G1").Value = Array("Onglet", "Signature", "Du", "Au", "Statut")
    r = 1
    For Each ws In ThisWorkbook.Worksheets
        If EstOngletSource(ws) Then
            r = r + 1: nb = nb + 1
            DatesOnglet ws, d1, d2
            wsL.Cells(r, 3).NumberFormat = "@": wsL.Cells(r, 4).NumberFormat = "@"
            wsL.Cells(r, 3).Value = ws.Name
            wsL.Cells(r, 4).Value = Signature(ws)
            If d1 > 0 Then wsL.Cells(r, 5).Value = CDate(d1): wsL.Cells(r, 6).Value = CDate(d2)
            wsL.Cells(r, 7).Value = "intégré"
        End If
    Next ws
    EcrireEtat "nb_onglets", CStr(nb)
    EcrireEtat "sig_abonnements", SignatureAbonnements()
End Sub
