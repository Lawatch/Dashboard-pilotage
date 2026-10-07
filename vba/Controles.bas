'==============================================================================
' CONTROLES : tout ce qui est à corriger ou à vérifier, sur la période de contrôle
' (PARAMETRES, 28 derniers jours de données par défaut)
'
'   1. Prestations sans abonnement (saisie du n° ou clic sur la suggestion)
'   2. Abonnements sans prestation
'   3. Prestations non affectées (agent « A SUPPRIMER », tournée de remplacement...)
'   4. Pointages à vérifier (début sans fin, durée anormale)
'   5. Abonnements en double ou à 0 EUR
'   6. Onglets jours (doublons, jours manquants, format non reconnu)
'   7. Chantiers aux coordonnées GPS absentes ou hors zone
'   8. Rattachements manuels invalides
'==============================================================================
Option Explicit

Private Const BLEU As Long = 6240799            ' RGB(31, 58, 95)
Private Const L_RESUME As Long = 5              ' 1re ligne du résumé
Private Const NB_SECTIONS As Long = 8
Private Const MAX_LIGNES As Long = 300

Private secTitre(1 To NB_SECTIONS) As String, secAide(1 To NB_SECTIONS) As String
Private secNb(1 To NB_SECTIONS) As Long, secEnjeu(1 To NB_SECTIONS) As String, secLigne(1 To NB_SECTIONS) As Long
' Suggestions de rattachement
Private sPret As Boolean, sNomAbo() As String, sRefAbo() As String
Private sLa() As Double, sLo() As Double, sAbo() As String, sNom() As String, nSite As Long

Public Sub EcrireControles(ByVal alertes As String)
    Dim ws As Worksheet, wsD As Worksheet, v As Variant, a As Variant, lastR As Long, r As Long, i As Long
    Dim dMax As Long, refDeb As Long, evts As Boolean, txtRes As String, nbTot As Long

    Set ws = ThisWorkbook.Worksheets(F_CTRL)
    Set wsD = ThisWorkbook.Worksheets(F_DATA)
    evts = Application.EnableEvents
    Application.EnableEvents = False
    dMax = DernierJour(): refDeb = dMax - JoursReference() + 1

    lastR = DerniereLigne(wsD)
    If lastR >= 2 Then v = wsD.Range(wsD.Cells(2, 1), wsD.Cells(lastR, NB_COLS)).Value Else v = Empty
    lastR = DerniereLigne(ThisWorkbook.Worksheets(F_ABOS))
    If lastR >= 2 Then a = ThisWorkbook.Worksheets(F_ABOS).Range("A2:P" & lastR).Value Else a = Empty

    EffacerZone ws.Range("A3:Z" & Application.WorksheetFunction.Max(200, DerniereLigne(ws) + 5))
    sPret = False
    ws.Range("B2").Value = "Période de contrôle : " & IIf(dMax > 0, TxtPeriode(refDeb, dMax), "aucune donnée") & _
        " (" & JoursReference() & " derniers jours de données, réglable dans PARAMETRES)."
    FormatAide ws.Range("B2")

    ' sections, écrites sous le résumé
    r = L_RESUME + NB_SECTIONS + 3
    r = SectionSansAbo(ws, r, v, a, refDeb, dMax, 1)
    r = SectionAbosSansPresta(ws, r, a, 2)
    r = SectionFictifs(ws, r, v, refDeb, dMax, 3)
    r = SectionPointages(ws, r, v, refDeb, dMax, 4)
    r = SectionDoublonsAbos(ws, r, a, 5)
    r = SectionOnglets(ws, r, v, alertes, 6)
    r = SectionGPS(ws, r, v, refDeb, dMax, 7)
    r = SectionManuels(ws, r, 8)

    ' résumé
    ws.Range("B4").Value = "RÉSUMÉ"
    FormatTitre ws.Range("B4:J4"), BLEU
    For i = 1 To NB_SECTIONS
        ws.Cells(L_RESUME + i - 1, 2).Value = secNb(i)
        ws.Cells(L_RESUME + i - 1, 2).HorizontalAlignment = xlCenter
        ws.Cells(L_RESUME + i - 1, 2).Font.Bold = True
        ws.Cells(L_RESUME + i - 1, 2).Font.Color = IIf(secNb(i) > 0, RGB(198, 40, 40), RGB(46, 125, 50))
        Lien ws.Cells(L_RESUME + i - 1, 3), "'" & F_CTRL & "'!B" & secLigne(i), secTitre(i)
        ws.Cells(L_RESUME + i - 1, 6).Value = secEnjeu(i)
        If secNb(i) > 0 And i <= 4 Then nbTot = nbTot + 1
    Next i
    FormatLignes ws.Range(ws.Cells(L_RESUME, 2), ws.Cells(L_RESUME + NB_SECTIONS - 1, 10))

    ' texte d'alerte du DASHBOARD
    If secNb(2) > 0 Then txtRes = txtRes & secNb(2) & " abonnement(s) sans prestation (" & secEnjeu(2) & ")"
    If secNb(1) > 0 Then txtRes = txtRes & IIf(txtRes <> "", "  " & ChrW(183) & "  ", "") & secNb(1) & " chantier(s) sans abonnement"
    If secNb(3) > 0 Then txtRes = txtRes & IIf(txtRes <> "", "  " & ChrW(183) & "  ", "") & secEnjeu(3) & " non affectées"
    If secNb(6) > 0 Then txtRes = txtRes & IIf(txtRes <> "", "  " & ChrW(183) & "  ", "") & "onglets jours à vérifier"
    EcrireEtat "nb_anomalies", CStr(secNb(1) + secNb(2) + secNb(3) + secNb(6))
    EcrireEtat "txt_anomalies", txtRes
    Application.EnableEvents = evts
End Sub

' Titre + aide + en-têtes d'une section ; renvoie la ligne des en-têtes
Private Function DebutSection(ws As Worksheet, ByVal r As Long, ByVal s As Long, entetes As Variant) As Long
    Dim n As Long
    secLigne(s) = r
    ws.Cells(r, 2).Value = s & ". " & UCase$(secTitre(s)) & "  (" & secNb(s) & ")"
    FormatTitre ws.Range(ws.Cells(r, 2), ws.Cells(r, 10)), IIf(secNb(s) > 0, BLEU, RGB(46, 125, 50))
    ws.Cells(r + 1, 2).Value = secAide(s)
    FormatAide ws.Cells(r + 1, 2)
    If secNb(s) = 0 Then DebutSection = 0: Exit Function
    n = UBound(entetes) - LBound(entetes) + 1
    ws.Range(ws.Cells(r + 2, 2), ws.Cells(r + 2, 1 + n)).Value = entetes
    FormatEntete ws.Range(ws.Cells(r + 2, 2), ws.Cells(r + 2, 1 + n))
    ws.Rows(r + 2).RowHeight = 30
    DebutSection = r + 2
End Function

Private Function FinSection(ws As Worksheet, ByVal r As Long, ByVal nLignes As Long, ByVal nTotal As Long) As Long
    If nTotal > nLignes Then
        ws.Cells(r + nLignes + 1, 2).Value = "... et " & (nTotal - nLignes) & " autre(s) ligne(s)."
        FormatAide ws.Cells(r + nLignes + 1, 2)
        nLignes = nLignes + 1
    End If
    If nLignes > 0 Then
        FormatLignes ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + nLignes, 10))
        ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + nLignes, 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(r + 1, 7), ws.Cells(r + nLignes, 8)).HorizontalAlignment = xlCenter
    End If
    FinSection = r + nLignes + 3
End Function

'------------------------------------------------------------------------------
' 1. Prestations sans abonnement, regroupées par projet / chantier
'------------------------------------------------------------------------------
Private Function SectionSansAbo(ws As Worksheet, ByVal r As Long, v As Variant, a As Variant, _
                                ByVal refDeb As Long, ByVal dMax As Long, ByVal s As Long) As Long
    Dim grp As New Collection, n As Long, i As Long, k As Long, d As Long, cle As String, h As Long, ag As String
    Dim gP() As String, gC() As String, gN() As Long, gH() As Double, gA() As String, gD() As Long, gLa() As Double, gLo() As Double
    Dim ord() As Long, vals() As Double, arr() As Variant, nL As Long, totH As Double

    secTitre(s) = "Prestations sans abonnement"
    secAide(s) = "Chantiers planifiés qui ne correspondent à aucun abonnement actif : prestation non facturée, ou projet à recoller. " & _
                 "Saisir le n" & ChrW(176) & " d'abonnement en colonne J, ou cliquer sur la suggestion : le rattachement est enregistré dans PARAMETRES."
    If Not IsEmpty(v) Then
        ReDim gP(1 To UBound(v, 1)): ReDim gC(1 To UBound(v, 1)): ReDim gN(1 To UBound(v, 1)): ReDim gH(1 To UBound(v, 1))
        ReDim gA(1 To UBound(v, 1)): ReDim gD(1 To UBound(v, 1)): ReDim gLa(1 To UBound(v, 1)): ReDim gLo(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            cle = txt(v(i, D_CLE))
            d = SerialDate(v(i, D_DATE))
            If (Left$(cle, 2) = "P:" Or Left$(cle, 2) = "C:") And d >= refDeb And d <= dMax Then
                k = Idx(grp, cle)
                If k = 0 Then
                    n = n + 1: k = n: Ajout grp, cle, k
                    gP(k) = txt(v(i, D_PROJET)): gC(k) = txt(v(i, D_CHANTIER))
                    If gC(k) = "" Then gC(k) = "(sans libellé)"
                End If
                gN(k) = gN(k) + 1
                gH(k) = gH(k) + Nombre0(v(i, D_DUREEP)): totH = totH + Nombre0(v(i, D_DUREEP))
                If d > gD(k) Then gD(k) = d
                ag = Trim$(txt(v(i, D_AGENT)))
                If InStr(1, "|" & gA(k) & "|", "|" & ag & "|") = 0 Then gA(k) = gA(k) & IIf(gA(k) = "", "", "|") & ag
                If gLa(k) = 0 And Nombre0(v(i, D_LAT)) <> 0 Then gLa(k) = Nombre0(v(i, D_LAT)): gLo(k) = Nombre0(v(i, D_LON))
            End If
        Next i
    End If
    secNb(s) = n
    secEnjeu(s) = Format(totH, "0") & " h prévues"
    h = DebutSection(ws, r, s, Array("Projet", "Chantier", "Prestations", "Heures prévues", "Agents", "Dernière", "", _
                                     "Suggestion (cliquer pour accepter)", "N" & ChrW(176) & " abonnement à rattacher"))
    If h = 0 Then SectionSansAbo = r + 3: Exit Function
    ReDim ord(1 To n): ReDim vals(1 To n)
    For k = 1 To n: ord(k) = k: vals(k) = gN(k) + gH(k) / 10000: Next k
    TriDesc vals, ord, 1, n
    nL = IIf(n > MAX_LIGNES, MAX_LIGNES, n)
    ReDim arr(1 To nL, 1 To 9)
    For i = 1 To nL
        k = ord(i)
        arr(i, 1) = gP(k): arr(i, 2) = gC(k): arr(i, 3) = gN(k): arr(i, 4) = Round(gH(k), 2)
        arr(i, 5) = Replace(gA(k), "|", ", "): arr(i, 6) = CDate(gD(k)): arr(i, 7) = ""
        arr(i, 8) = Suggestion(gC(k), gLa(k), gLo(k), v, a)
        arr(i, 9) = ""
    Next i
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 10)).Value = arr
    ws.Range(ws.Cells(h + 1, 7), ws.Cells(h + nL, 7)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(h + 1, 1), ws.Cells(h + nL, 1)).Value = "SANSABO"
    With ws.Range(ws.Cells(h + 1, 10), ws.Cells(h + nL, 10))
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .NumberFormat = "0"
    End With
    For i = 1 To nL
        If arr(i, 8) <> "" Then
            Cliquable ws.Cells(h + i, 9), CStr(arr(i, 8))
            ws.Cells(h + i, 9).Font.Size = 9
        End If
    Next i
    SectionSansAbo = FinSection(ws, h, nL, n)
End Function

' Abonnement probable pour un chantier sans abonnement : mêmes mots dans le nom, sinon même adresse GPS
Private Function Suggestion(ByVal chantier As String, ByVal la As Double, ByVal lo As Double, v As Variant, a As Variant) As String
    Dim i As Long, best As Double, sc As Double, iBest As Long, dBest As Double, dist As Double, nbEgaux As Long, c As String
    If IsEmpty(a) Then Exit Function
    PreparerSuggestions v, a
    c = NormC(chantier)
    For i = 1 To UBound(a, 1)
        sc = Ressemblance(c, sNomAbo(i))
        If Ressemblance(c, sRefAbo(i)) > sc Then sc = Ressemblance(c, sRefAbo(i))
        If sc > 0 And Nombre0(a(i, 13)) = 0 Then sc = sc + 0.05            ' de préférence un abonnement sans prestation
        If sc > best + 0.001 Then
            best = sc: iBest = i: nbEgaux = 1
        ElseIf Abs(sc - best) <= 0.001 And sc > 0 Then
            nbEgaux = nbEgaux + 1
        End If
    Next i
    If best >= 0.6 Then
        Suggestion = txt(a(iBest, 1)) & " " & ChrW(183) & " " & txt(a(iBest, 2)) & _
                     IIf(txt(a(iBest, 10)) <> "" And txt(a(iBest, 10)) <> txt(a(iBest, 2)), " (" & txt(a(iBest, 10)) & ")", "") & _
                     IIf(nbEgaux > 1, "  [" & nbEgaux & " possibles]", "")
        Exit Function
    End If
    ' même lieu qu'un chantier déjà rattaché (moins de 100 m)
    If la = 0 Then Exit Function
    dBest = 0.1: iBest = 0
    For i = 1 To nSite
        dist = DistKm(la, lo, sLa(i), sLo(i))
        If dist < dBest Then dBest = dist: iBest = i
    Next i
    If iBest > 0 Then Suggestion = sAbo(iBest) & " " & ChrW(183) & " " & sNom(iBest) & " (à " & Format(dBest * 1000, "0") & " m)"
End Function

' Noms d'abonnements normalisés et chantiers déjà rattachés (une fois par mise à jour)
Private Sub PreparerSuggestions(v As Variant, a As Variant)
    Dim i As Long, vus As New Collection, cle As String
    If sPret Then Exit Sub
    ReDim sNomAbo(1 To UBound(a, 1)): ReDim sRefAbo(1 To UBound(a, 1))
    For i = 1 To UBound(a, 1)
        sNomAbo(i) = NormC(txt(a(i, 2))): sRefAbo(i) = NormC(txt(a(i, 10)))
    Next i
    nSite = 0
    If Not IsEmpty(v) Then
        ReDim sLa(1 To UBound(v, 1)): ReDim sLo(1 To UBound(v, 1)): ReDim sAbo(1 To UBound(v, 1)): ReDim sNom(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            cle = txt(v(i, D_CLE))
            If cle <> "" And Left$(cle, 2) <> "P:" And Left$(cle, 2) <> "C:" And Nombre0(v(i, D_LAT)) <> 0 Then
                If AjoutUnique(vus, Format(Nombre0(v(i, D_LAT)), "0.0000") & Format(Nombre0(v(i, D_LON)), "0.0000") & cle) Then
                    nSite = nSite + 1
                    sLa(nSite) = Nombre0(v(i, D_LAT)): sLo(nSite) = Nombre0(v(i, D_LON))
                    sAbo(nSite) = txt(v(i, D_ABO)): sNom(nSite) = txt(v(i, D_CHANTIER))
                End If
            End If
        Next i
    End If
    sPret = True
End Sub

' Ressemblance de deux noms normalisés (0 à 1) : mots significatifs en commun,
' rapportés au plus long des deux noms (« LES CEDRES » ne ressemble pas à « CLINIQUE DES CEDRES »)
Private Function Ressemblance(ByVal n1 As String, ByVal n2 As String) As Double
    Dim m() As String, i As Long, t1 As Long, t2 As Long, ok As Long, b As String
    If n1 = "" Or n2 = "" Then Exit Function
    m = Split(n2, " ")
    For i = 0 To UBound(m)
        If Len(m(i)) >= 1 And Not MotVide(m(i)) Then t2 = t2 + 1
    Next i
    b = " " & n2 & " "
    m = Split(n1, " ")
    For i = 0 To UBound(m)
        If Len(m(i)) >= 1 And Not MotVide(m(i)) Then
            t1 = t1 + 1
            If InStr(b, " " & m(i) & " ") > 0 Then ok = ok + 1
        End If
    Next i
    If t1 = 0 Or t2 = 0 Then Exit Function
    Ressemblance = ok / IIf(t1 > t2, t1, t2)
End Function

Private Function MotVide(ByVal w As String) As Boolean
    MotVide = InStr("|LE|LA|LES|DE|DU|DES|D|L|ET|A|AU|AUX|RUE|AVENUE|AV|BD|BOULEVARD|COURS|PLACE|CHEMIN|GESTION|CONTAINERS|" & _
                    "CONTAINER|CONTENEURS|ENTRETIEN|PC|GARAGES|GARAGE|PARTIES|COMMUNES|RESIDENCE|SDC|BAT|BATIMENT|", "|" & w & "|") > 0
End Function

'------------------------------------------------------------------------------
' 2. Abonnements sans prestation
'------------------------------------------------------------------------------
Private Function SectionAbosSansPresta(ws As Worksheet, ByVal r As Long, a As Variant, ByVal s As Long) As Long
    Dim i As Long, n As Long, ord() As Long, vals() As Double, arr() As Variant, h As Long, k As Long, tot As Double, nL As Long
    secTitre(s) = "Abonnements sans prestation"
    secAide(s) = "Abonnements actifs (facturés) sans aucune prestation sur la période de contrôle : prestation oubliée au planning, " & _
                 "projet différent (voir section 1), ou abonnement à résilier. « Dernière » = dernière prestation connue."
    If Not IsEmpty(a) Then
        ReDim ord(1 To UBound(a, 1)): ReDim vals(1 To UBound(a, 1))
        For i = 1 To UBound(a, 1)
            If Nombre0(a(i, 13)) = 0 Then
                n = n + 1: ord(n) = i: vals(n) = Nombre0(a(i, 6)): tot = tot + Nombre0(a(i, 6))
            End If
        Next i
    End If
    secNb(s) = n
    secEnjeu(s) = Euros(tot) & " HT / mois"
    h = DebutSection(ws, r, s, Array("N" & ChrW(176) & " abonnement", "Chantier", "CA mensuel", "Fréquence", "Client facturation", _
                                     "Projet", "Dernière prestation", "Adresse"))
    If h = 0 Then SectionAbosSansPresta = r + 3: Exit Function
    TriDesc vals, ord, 1, n
    nL = IIf(n > MAX_LIGNES, MAX_LIGNES, n)
    ReDim arr(1 To nL, 1 To 8)
    For i = 1 To nL
        k = ord(i)
        arr(i, 1) = a(k, 1): arr(i, 2) = txt(a(k, 2)): arr(i, 3) = Nombre0(a(k, 6)): arr(i, 4) = txt(a(k, 4))
        arr(i, 5) = txt(a(k, 3)): arr(i, 6) = txt(a(k, 7))
        If EstUneDate(a(k, 12)) Then arr(i, 7) = CDate(SerialDate(a(k, 12))) Else arr(i, 7) = "jamais"
        arr(i, 8) = txt(a(k, 9))
    Next i
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 9)).Value = arr
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 2)).NumberFormat = "0"
    ws.Range(ws.Cells(h + 1, 4), ws.Cells(h + nL, 4)).NumberFormat = "#,##0 " & ChrW(8364)
    ws.Range(ws.Cells(h + 1, 8), ws.Cells(h + nL, 8)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(h + 1, 8), ws.Cells(h + nL, 8)).HorizontalAlignment = xlRight
    SectionAbosSansPresta = FinSection(ws, h, nL, n)
End Function

'------------------------------------------------------------------------------
' 3. Prestations non affectées (agent fictif)
'------------------------------------------------------------------------------
Private Function SectionFictifs(ws As Worksheet, ByVal r As Long, v As Variant, ByVal refDeb As Long, ByVal dMax As Long, _
                                ByVal s As Long) As Long
    Dim grp As New Collection, n As Long, i As Long, k As Long, d As Long, cle As String, h As Long, nTot As Long, hTot As Double
    Dim gAg() As String, gC() As String, gN() As Long, gH() As Double, gJ() As String, gD() As Long, gAbo() As String
    Dim ord() As Long, vals() As Double, arr() As Variant, nL As Long
    secTitre(s) = "Prestations non affectées"
    secAide(s) = "Prestations planifiées sur un agent fictif (« A SUPPRIMER », tournée de remplacement...) : personne n'y est réellement affecté."
    If Not IsEmpty(v) Then
        ReDim gAg(1 To UBound(v, 1)): ReDim gC(1 To UBound(v, 1)): ReDim gN(1 To UBound(v, 1)): ReDim gH(1 To UBound(v, 1))
        ReDim gJ(1 To UBound(v, 1)): ReDim gD(1 To UBound(v, 1)): ReDim gAbo(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            d = SerialDate(v(i, D_DATE))
            If d >= refDeb And d <= dMax And EstAgentFictif(txt(v(i, D_AGENT))) Then
                cle = Trim$(txt(v(i, D_AGENT))) & "|" & txt(v(i, D_CLE)) & "|" & txt(v(i, D_LIBELLE))
                k = Idx(grp, cle)
                If k = 0 Then
                    n = n + 1: k = n: Ajout grp, cle, k
                    gAg(k) = Trim$(txt(v(i, D_AGENT))): gC(k) = txt(v(i, D_LIBELLE))
                    If gC(k) = "" Then gC(k) = txt(v(i, D_CHANTIER))
                    gAbo(k) = txt(v(i, D_ABO))
                End If
                gN(k) = gN(k) + 1: nTot = nTot + 1
                gH(k) = gH(k) + Nombre0(v(i, D_DUREEP)): hTot = hTot + Nombre0(v(i, D_DUREEP))
                If InStr(gJ(k), Format(d, "dd/mm")) = 0 Then gJ(k) = gJ(k) & IIf(gJ(k) = "", "", ", ") & Format(d, "dd/mm")
                If d > gD(k) Then gD(k) = d
            End If
        Next i
    End If
    secNb(s) = n
    secEnjeu(s) = nTot & " prestations (" & Format(hTot, "0") & " h)"
    h = DebutSection(ws, r, s, Array("Agent planning", "Prestation", "Nombre", "Heures prévues", "Jours", "Dernière", _
                                     "N" & ChrW(176) & " abonnement"))
    If h = 0 Then SectionFictifs = r + 3: Exit Function
    ReDim ord(1 To n): ReDim vals(1 To n)
    For k = 1 To n: ord(k) = k: vals(k) = gN(k) + gH(k) / 10000: Next k
    TriDesc vals, ord, 1, n
    nL = IIf(n > MAX_LIGNES, MAX_LIGNES, n)
    ReDim arr(1 To nL, 1 To 7)
    For i = 1 To nL
        k = ord(i)
        arr(i, 1) = gAg(k): arr(i, 2) = gC(k): arr(i, 3) = gN(k): arr(i, 4) = Round(gH(k), 2)
        arr(i, 5) = gJ(k): arr(i, 6) = CDate(gD(k)): arr(i, 7) = IIf(gAbo(k) = "", "sans abonnement", gAbo(k))
    Next i
    ws.Range(ws.Cells(h + 1, 6), ws.Cells(h + nL, 6)).NumberFormat = "@"
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 8)).Value = arr
    ws.Range(ws.Cells(h + 1, 7), ws.Cells(h + nL, 7)).NumberFormat = "dd/mm/yyyy"
    SectionFictifs = FinSection(ws, h, nL, n)
End Function

'------------------------------------------------------------------------------
' 4. Pointages à vérifier
'------------------------------------------------------------------------------
Private Function SectionPointages(ws As Worksheet, ByVal r As Long, v As Variant, ByVal refDeb As Long, ByVal dMax As Long, _
                                  ByVal s As Long) As Long
    Dim i As Long, n As Long, d As Long, st As String, cles() As String, ix() As Long, arr() As Variant, h As Long, k As Long, nL As Long
    secTitre(s) = "Pointages à vérifier"
    secAide(s) = "Début pointé sans fin ; durée anormale (moins du quart, ou plus du triple du prévu) ; clics simultanés " & _
                 "(plusieurs prestations pointées aux mêmes minutes : clics faits d'un coup, pas sur place). Conteneurs exclus."
    If Not IsEmpty(v) Then
        ReDim cles(1 To UBound(v, 1)): ReDim ix(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            d = SerialDate(v(i, D_DATE))
            st = txt(v(i, D_STATUT))
            If d >= refDeb And d <= dMax And txt(v(i, D_COMPTE)) <> "0" And _
               (st = "Début seul" Or st = "Fin seule" Or st = "Trop courte" Or st = "Trop longue" Or st = "Clics simultanés") Then
                n = n + 1: ix(n) = i
                cles(n) = Format(99999 - d, "00000") & "|" & txt(v(i, D_AGENT)) & "|" & Format(Minutes(v(i, D_DEBR)), "0000")
            End If
        Next i
    End If
    secNb(s) = n
    secEnjeu(s) = n & " prestation(s)"
    h = DebutSection(ws, r, s, Array("Date", "Prestation", "Clic début", "Clic fin", "Agent", "Durée pointée (h)", _
                                     "Durée prévue (h)", "Constat"))
    If h = 0 Then SectionPointages = r + 3: Exit Function
    TriTexte cles, ix, 1, n
    nL = IIf(n > MAX_LIGNES, MAX_LIGNES, n)
    ReDim arr(1 To nL, 1 To 8)
    For k = 1 To nL
        i = ix(k)
        arr(k, 1) = CDate(SerialDate(v(i, D_DATE))): arr(k, 2) = txt(v(i, D_LIBELLE))
        arr(k, 3) = HHMM(v(i, D_DEBR)): arr(k, 4) = HHMM(v(i, D_FINR)): arr(k, 5) = txt(v(i, D_AGENT))
        arr(k, 6) = IIf(IsEmpty(v(i, D_DUREER)), "", v(i, D_DUREER)): arr(k, 7) = Nombre0(v(i, D_DUREEP))
        arr(k, 8) = txt(v(i, D_STATUT))
    Next k
    ws.Range(ws.Cells(h + 1, 4), ws.Cells(h + nL, 5)).NumberFormat = "@"
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 9)).Value = arr
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 2)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(h + 1, 7), ws.Cells(h + nL, 8)).NumberFormat = "0.00"
    SectionPointages = FinSection(ws, h, nL, n)
End Function

'------------------------------------------------------------------------------
' 5. Abonnements en double (même projet ou même chantier, même montant) ou à 0 EUR
'------------------------------------------------------------------------------
Private Function SectionDoublonsAbos(ws As Worksheet, ByVal r As Long, a As Variant, ByVal s As Long) As Long
    Dim i As Long, j As Long, n As Long, motif() As String, arr() As Variant, h As Long, k As Long, nL As Long
    Dim cles() As String, ix() As Long, grp As New Collection, cle As String, cle2 As String, nbG As New Collection
    secTitre(s) = "Abonnements en double ou à 0 " & ChrW(8364)
    secAide(s) = "Même projet ou même chantier avec le même montant (facturé deux fois ?), ou montant nul."
    If Not IsEmpty(a) Then
        ReDim motif(1 To UBound(a, 1)): ReDim cles(1 To UBound(a, 1)): ReDim ix(1 To UBound(a, 1))
        For i = 1 To UBound(a, 1)
            If Nombre0(a(i, 5)) > 0 Then
                cle = "P" & txt(a(i, 7)) & "|" & Format(Nombre0(a(i, 5)), "0.00")
                cle2 = "C" & Norm(txt(a(i, 2))) & "|" & Format(Nombre0(a(i, 5)), "0.00")
                If txt(a(i, 7)) <> "" Then AjouterUn nbG, cle
                If Norm(txt(a(i, 2))) <> "" Then AjouterUn nbG, cle2
            End If
        Next i
        For i = 1 To UBound(a, 1)
            If Nombre0(a(i, 5)) = 0 Then
                motif(i) = "Montant à 0"
            Else
                cle = "P" & txt(a(i, 7)) & "|" & Format(Nombre0(a(i, 5)), "0.00")
                cle2 = "C" & Norm(txt(a(i, 2))) & "|" & Format(Nombre0(a(i, 5)), "0.00")
                If txt(a(i, 7)) <> "" And Idx(nbG, cle) > 1 Then
                    motif(i) = "Même projet, même montant"
                ElseIf Idx(nbG, cle2) > 1 Then
                    motif(i) = "Même chantier, même montant"
                End If
            End If
            If motif(i) <> "" Then
                n = n + 1: ix(n) = i
                cles(n) = IIf(motif(i) = "Montant à 0", "1", "0") & Norm(txt(a(i, 2))) & "|" & txt(a(i, 1))
            End If
        Next i
    End If
    secNb(s) = n
    secEnjeu(s) = n & " abonnement(s)"
    h = DebutSection(ws, r, s, Array("N" & ChrW(176) & " abonnement", "Chantier", "Montant HT", "Fréquence", "Client facturation", _
                                     "Projet", "Date début", "Constat"))
    If h = 0 Then SectionDoublonsAbos = r + 3: Exit Function
    TriTexte cles, ix, 1, n
    nL = IIf(n > MAX_LIGNES, MAX_LIGNES, n)
    ReDim arr(1 To nL, 1 To 8)
    For k = 1 To nL
        i = ix(k)
        arr(k, 1) = a(i, 1): arr(k, 2) = txt(a(i, 2)): arr(k, 3) = Nombre0(a(i, 5)): arr(k, 4) = txt(a(i, 4))
        arr(k, 5) = txt(a(i, 3)): arr(k, 6) = txt(a(i, 7)): arr(k, 7) = a(i, 15): arr(k, 8) = motif(i)
    Next k
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 9)).Value = arr
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + nL, 2)).NumberFormat = "0"
    ws.Range(ws.Cells(h + 1, 4), ws.Cells(h + nL, 4)).NumberFormat = "#,##0.00 " & ChrW(8364)
    ws.Range(ws.Cells(h + 1, 8), ws.Cells(h + nL, 8)).NumberFormat = "dd/mm/yyyy"
    SectionDoublonsAbos = FinSection(ws, h, nL, n)
End Function

Private Sub AjouterUn(c As Collection, ByVal cle As String)
    Dim x As Long
    x = Idx(c, cle)
    If x > 0 Then c.Remove "k" & cle
    c.Add x + 1, "k" & cle
End Sub

'------------------------------------------------------------------------------
' 6. Onglets jours
'------------------------------------------------------------------------------
Private Function SectionOnglets(ws As Worksheet, ByVal r As Long, v As Variant, ByVal alertes As String, ByVal s As Long) As Long
    Dim lignes() As Variant, n As Long, wsX As Worksheet, h As Long, i As Long, d As Long, dMin As Long, dMax As Long
    Dim jours As New Collection, wsL As Worksheet, lastR As Long
    ReDim lignes(1 To 600, 1 To 2)
    secTitre(s) = "Onglets jours"
    secAide(s) = "Onglets mis de côté (à supprimer), jours présents deux fois, jours ouvrés sans données, onglets illisibles."
    For Each wsX In ThisWorkbook.Worksheets
        If EstOngletEcarte(wsX.Name) And n < 600 Then
            n = n + 1: lignes(n, 1) = wsX.Name
            lignes(n, 2) = "Mis de côté (doublon ou ancienne version) : n'est pas pris en compte, peut être supprimé."
        ElseIf Not EstOngletSysteme(wsX.Name) And Not EstOngletEcarte(wsX.Name) And n < 600 Then
            If wsX.Name Like "######*" And Not EstOngletSource(wsX) Then
                n = n + 1: lignes(n, 1) = wsX.Name
                lignes(n, 2) = "Format non reconnu : la colonne G doit contenir « Nom Collaborateur » (coller l'extract complet, en-têtes compris)."
            End If
        End If
    Next wsX
    Set wsL = ThisWorkbook.Worksheets(F_LISTES)
    lastR = DerniereLigne(wsL)
    For i = 2 To lastR
        If EstUneDate(wsL.Cells(i, 12).Value) And n < 600 Then
            n = n + 1: lignes(n, 1) = txt(wsL.Cells(i, 13).Value)
            lignes(n, 2) = "Contient aussi le " & Format(wsL.Cells(i, 12).Value, "dd/mm/yyyy") & ", déjà présent dans un autre onglet : ignoré pour ce jour."
        End If
    Next i
    ' jours ouvrés (lundi à vendredi) sans données entre le premier et le dernier jour
    If Not IsEmpty(v) Then
        For i = 1 To UBound(v, 1)
            d = SerialDate(v(i, D_DATE))
            If d > 0 Then
                AjoutUnique jours, CStr(d)
                If dMin = 0 Or d < dMin Then dMin = d
                If d > dMax Then dMax = d
            End If
        Next i
        For d = dMin To dMax
            If Weekday(d, vbMonday) <= 5 And Not Existe(jours, CStr(d)) And n < 600 Then
                n = n + 1: lignes(n, 1) = JJMMAA(d)
                lignes(n, 2) = "Aucun onglet pour ce jour ouvré (" & Format(d, "dddd dd/mm/yyyy") & ")."
            End If
        Next d
    End If
    secNb(s) = n
    secEnjeu(s) = n & " point(s)"
    h = DebutSection(ws, r, s, Array("Onglet / jour", "Constat"))
    If h = 0 Then SectionOnglets = r + 3: Exit Function
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + n, 2)).NumberFormat = "@"
    For i = 1 To n
        ws.Cells(h + i, 2).Value = lignes(i, 1)
        ws.Cells(h + i, 3).Value = lignes(i, 2)
    Next i
    SectionOnglets = FinSection(ws, h, n, n)
End Function

'------------------------------------------------------------------------------
' 7. Coordonnées GPS des chantiers absentes ou hors zone
'------------------------------------------------------------------------------
Private Function SectionGPS(ws As Worksheet, ByVal r As Long, v As Variant, ByVal refDeb As Long, ByVal dMax As Long, _
                            ByVal s As Long) As Long
    Dim wsP As Worksheet, agLa As Double, agLo As Double, rayon As Double, i As Long, n As Long, k As Long, d As Long
    Dim grp As New Collection, cle As String, arr() As Variant, h As Long, la As Double, lo As Double, ag As String, nL As Long
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)
    agLa = Nombre0(wsP.Range("H11").Value): agLo = Nombre0(wsP.Range("I11").Value)
    rayon = Nombre0(wsP.Range(P_RAYON).Value): If rayon <= 0 Then rayon = 60
    secTitre(s) = "Coordonnées GPS des chantiers"
    secAide(s) = "Chantiers sans coordonnées ou à plus de " & rayon & " km de l'agence : absents de la carte et des km. À corriger dans l'outil (fiche projet)."
    If Not IsEmpty(v) Then
        ReDim arr(1 To MAX_LIGNES, 1 To 6)
        For i = 1 To UBound(v, 1)
            d = SerialDate(v(i, D_DATE))
            If d >= refDeb And d <= dMax Then
                la = Nombre0(v(i, D_LAT)): lo = Nombre0(v(i, D_LON))
                If la = 0 Or lo = 0 Or (agLa <> 0 And DistKm(agLa, agLo, la, lo) > rayon) Then
                    cle = txt(v(i, D_PROJET)) & "|" & txt(v(i, D_LIBELLE))
                    k = Idx(grp, cle)
                    If k = 0 And n < MAX_LIGNES Then
                        n = n + 1: k = n: Ajout grp, cle, k
                        arr(k, 1) = txt(v(i, D_PROJET)): arr(k, 2) = txt(v(i, D_LIBELLE))
                        arr(k, 3) = IIf(la = 0, "absente", la): arr(k, 4) = IIf(lo = 0, "absente", lo): arr(k, 5) = 0
                    End If
                    If k > 0 Then
                        arr(k, 5) = arr(k, 5) + 1
                        ag = Trim$(txt(v(i, D_AGENT)))
                        If InStr(1, ", " & arr(k, 6) & ",", ", " & ag & ",") = 0 Then arr(k, 6) = arr(k, 6) & IIf(arr(k, 6) = "", "", ", ") & ag
                    End If
                End If
            End If
        Next i
    End If
    secNb(s) = n
    secEnjeu(s) = n & " chantier(s)"
    h = DebutSection(ws, r, s, Array("Projet", "Prestation", "Latitude", "Longitude", "Prestations", "Agents"))
    If h = 0 Then SectionGPS = r + 3: Exit Function
    ws.Range(ws.Cells(h + 1, 2), ws.Cells(h + n, 7)).Value = arr
    SectionGPS = FinSection(ws, h, n, n)
End Function

'------------------------------------------------------------------------------
' 8. Rattachements manuels invalides
'------------------------------------------------------------------------------
Private Function SectionManuels(ws As Worksheet, ByVal r As Long, ByVal s As Long) As Long
    Dim wsP As Worksheet, lastR As Long, i As Long, n As Long, h As Long, arr() As Variant
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)
    secTitre(s) = "Rattachements manuels invalides"
    secAide(s) = "Rattachements saisis dans PARAMETRES vers un n" & ChrW(176) & " d'abonnement inconnu ou terminé : corriger ou effacer la ligne."
    lastR = DerniereLigne(wsP)
    ReDim arr(1 To Application.WorksheetFunction.Max(1, lastR), 1 To 3)
    For i = P_MAN1 To lastR
        If txt(wsP.Cells(i, 12).Value) <> "" And txt(wsP.Cells(i, 14).Value) <> "OK" And txt(wsP.Cells(i, 14).Value) <> "" Then
            n = n + 1
            arr(n, 1) = txt(wsP.Cells(i, 12).Value): arr(n, 2) = txt(wsP.Cells(i, 13).Value): arr(n, 3) = txt(wsP.Cells(i, 14).Value)
        End If
    Next i
    secNb(s) = n
    secEnjeu(s) = n & " ligne(s)"
    h = DebutSection(ws, r, s, Array("Projet ou chantier", "N" & ChrW(176) & " saisi", "Constat"))
    If h = 0 Then SectionManuels = r + 3: Exit Function
    For i = 1 To n
        ws.Cells(h + i, 2).Value = arr(i, 1): ws.Cells(h + i, 3).Value = arr(i, 2): ws.Cells(h + i, 4).Value = arr(i, 3)
    Next i
    SectionManuels = FinSection(ws, h, n, n)
End Function

'------------------------------------------------------------------------------
' Clic sur une suggestion : le n° est recopié dans la colonne J puis rattaché
'------------------------------------------------------------------------------
Public Sub AccepterSuggestion(c As Range)
    Dim n As Variant
    n = Nombre(Split(txt(c.Value), " ")(0))
    If IsEmpty(n) Then Exit Sub
    Application.EnableEvents = False
    c.Worksheet.Cells(c.Row, 10).Value = n
    Application.EnableEvents = True
    MajRattachements
End Sub
