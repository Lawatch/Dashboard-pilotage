'==============================================================================
' DASHBOARD : indicateurs de la période, tableau par agent ou fiche d'un agent
'   Agent = « Tous » : indicateurs + tableau par agent + tableau par jour
'   Agent = un nom   : indicateurs + ses abonnements + ses prestations dans l'ordre des clics
' CA attribué à un agent = CA mensuel de chaque chantier x sa part des heures prévues
' sur la période de contrôle (PARAMETRES, 28 jours par défaut).
'==============================================================================
Option Explicit

Public Const CELL_AGENT As String = "C5"
Public Const CELL_DU As String = "H5"
Public Const CELL_AU As String = "J5"
Private Const Z_L1 As Long = 11                ' 1re ligne de la zone écrite par la macro
Private Const BLEU As Long = 6240799           ' RGB(31, 58, 95)

' Abonnements regroupés par chantier (clé = n° du 1er abonnement du projet)
Private aCle As Collection, aK() As String, aNom() As String, aNum() As String, aCli() As String, aFreq() As String
Private aMens() As Double, nA As Long
' Tournées (km pointés), relues à chaque actualisation
Private vT As Variant, nT As Long
' Heures prévues agent x chantier (période de contrôle) et CA mensuel attribué à chaque agent
Private colPaire As Collection, pAg() As String, pCle() As String, pH() As Double, nPaire As Long
Private caParAgent As Collection

'------------------------------------------------------------------------------
' Filtres
'------------------------------------------------------------------------------
Public Sub PeriodeParDefaut()
    ChoisirPeriode 1, False
End Sub

' 1 = dernier jour, 7, 30, 0 = tout
Public Sub ChoisirPeriode(ByVal nbJours As Long, Optional ByVal actualiser As Boolean = True)
    Dim dMax As Long, dMin As Long, evts As Boolean
    dMax = DernierJour(): dMin = PremierJour()
    If dMax = 0 Then Exit Sub
    evts = Application.EnableEvents
    Application.EnableEvents = False
    With ThisWorkbook.Worksheets(F_DASH)
        .Range(CELL_AU).Value = CDate(dMax)
        If nbJours <= 0 Then .Range(CELL_DU).Value = CDate(dMin) Else .Range(CELL_DU).Value = CDate(Application.WorksheetFunction.Max(dMin, dMax - nbJours + 1))
    End With
    Application.EnableEvents = evts
    If actualiser Then ActualiserDashboard
End Sub

Public Sub ChoisirAgent(ByVal nom As String)
    Dim evts As Boolean
    evts = Application.EnableEvents
    Application.EnableEvents = False
    ThisWorkbook.Worksheets(F_DASH).Range(CELL_AGENT).Value = nom
    Application.EnableEvents = evts
    ActualiserDashboard
    ThisWorkbook.Worksheets(F_DASH).Range("A1").Select
End Sub

'------------------------------------------------------------------------------
' Calcul et écriture
'------------------------------------------------------------------------------
Public Sub ActualiserDashboard()
    Dim ws As Worksheet, wsD As Worksheet, v As Variant, lastR As Long, i As Long, evts As Boolean, su As Boolean
    Dim agent As String, tous As Boolean, du As Long, au As Long, d As Long, refDeb As Long, dMax As Long
    Dim nP As Long, nOk As Long, hP As Double, hR As Double, chantiers As New Collection, nCh As Long
    Dim km As Double, nbJ As Long, ag As String, compte As Boolean, cle As String, dansPer As Boolean
    Dim hCle As New Collection, caAg As Double, caTot As Double, caServi As Double

    Set ws = ThisWorkbook.Worksheets(F_DASH)
    Set wsD = ThisWorkbook.Worksheets(F_DATA)
    evts = Application.EnableEvents: su = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    On Error GoTo Fin

    agent = Trim$(txt(ws.Range(CELL_AGENT).Value))
    If agent = "" Then agent = "Tous": ws.Range(CELL_AGENT).Value = "Tous"
    tous = (agent = "Tous")
    dMax = DernierJour()
    If Not EstUneDate(ws.Range(CELL_DU).Value) Or Not EstUneDate(ws.Range(CELL_AU).Value) Then ChoisirPeriode 1, False
    du = SerialDate(ws.Range(CELL_DU).Value): au = SerialDate(ws.Range(CELL_AU).Value)
    If du > au Then d = du: du = au: au = d
    refDeb = dMax - JoursReference() + 1

    ChargerAbos
    ChargerTournees
    EffacerZone ws.Range("A" & Z_L1 & ":X" & Application.WorksheetFunction.Max(Z_L1 + 50, DerniereLigne(ws) + 5))
    EcrireEtatDashboard ws

    lastR = DerniereLigne(wsD)
    If lastR < 2 Then
        For i = 0 To 6
            ws.Range(Split("B8,C8,E8,G8,I8,K8,M8", ",")(i)).Value = ""
            ws.Range(Split("B9,C9,E9,G9,I9,K9,M9", ",")(i)).Value = ""
        Next i
        ws.Range("B" & Z_L1).Value = "Aucune prestation : collez un extract de validations dans un nouvel onglet."
        GoTo Fin
    End If
    v = wsD.Range(wsD.Cells(2, 1), wsD.Cells(lastR, NB_COLS)).Value

    ' Heures prévues par chantier et par agent x chantier sur la période de contrôle (pour le CA)
    Set colPaire = New Collection: nPaire = 0
    ReDim pAg(1 To 1000): ReDim pCle(1 To 1000): ReDim pH(1 To 1000)
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, D_DATE))
        If d >= refDeb And d <= dMax Then
            ag = Trim$(txt(v(i, D_AGENT)))
            If Not EstAgentFictif(ag) Then
                cle = txt(v(i, D_CLE))
                AjouterNb hCle, cle, Nombre0(v(i, D_DUREEP))
                AjouterPaire ag, cle, Nombre0(v(i, D_DUREEP))
            End If
        End If
    Next i

    CalculerCAAgents hCle

    ' Indicateurs de la période
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, D_DATE))
        If d >= du And d <= au Then
            ag = Trim$(txt(v(i, D_AGENT)))
            compte = (txt(v(i, D_COMPTE)) <> "0")
            If (tous And compte) Or ag = agent Then
                nP = nP + 1
                If txt(v(i, D_POINTEE)) = "1" Then nOk = nOk + 1
                hP = hP + Nombre0(v(i, D_DUREEP))
                hR = hR + Nombre0(v(i, D_HRET))
                cle = txt(v(i, D_CLE))
                If cle <> "" Then
                    If AjoutUnique(chantiers, cle) Then
                        nCh = nCh + 1
                        If Idx(aCle, cle) > 0 Then caServi = caServi + aMens(Idx(aCle, cle))
                    End If
                End If
            End If
        End If
    Next i
    km = KmPeriode(IIf(tous, "", agent), du, au, nbJ, tous)

    ' Tuiles
    ws.Range("B8").Value = nP
    ws.Range("B9").Value = IIf(tous, "agents comptés", "sur la période")
    ws.Range("C8").Value = Pct(nOk, nP)
    ws.Range("C9").Value = nOk & " sur " & nP
    ws.Range("E8").Value = Round(hP, 1)
    ws.Range("E9").Value = "durées du planning"
    ws.Range("G8").Value = Round(hR, 1)
    ws.Range("G9").Value = "conteneurs : durée prévue"
    ws.Range("I8").Value = nCh
    ws.Range("I9").Value = "dont " & CompterAbos(chantiers) & " avec abonnement"
    If tous Then
        For i = 1 To nA: caTot = caTot + aMens(i): Next i
        ws.Range("K8").Value = Round(caTot, 0)
        ws.Range("K9").Value = "abonnements actifs (HT)"
    Else
        caAg = CAAgent(agent)
        ws.Range("K8").Value = Round(caAg, 0)
        ws.Range("K9").Value = "sa part, sur " & JoursReference() & " j"
    End If
    ws.Range("M8").Value = Round(km, 0)
    If nbJ > 0 Then ws.Range("M9").Value = Format(km / nbJ, "0") & " km / jour-agent" Else ws.Range("M9").Value = ""

    If tous Then
        EcrireParAgent ws, v, du, au
        EcrireParJour ws, v, du, au
    Else
        EcrireFicheAgent ws, v, agent, du, au, hCle
    End If
Fin:
    If Err.Number <> 0 Then ws.Range("B3").Value = ChrW(9888) & " Erreur d'affichage : " & Err.Description
    Application.EnableEvents = evts: Application.ScreenUpdating = su
End Sub

Private Sub EcrireEtatDashboard(ws As Worksheet)
    Dim dMin As Long, dMax As Long, t As String, nb As String, s As String
    dMin = PremierJour(): dMax = DernierJour()
    If dMax > 0 Then
        t = "Prestations du " & TxtPeriode(dMin, dMax) & " (" & LireEtat("nb_jours") & " jours)"
    Else
        t = "Aucune prestation"
    End If
    t = t & "   |   Abonnements : " & LireEtat("nb_abos") & " actifs"
    t = t & "   |   Mis à jour le " & LireEtat("maj")
    ws.Range("B2").Value = t
    nb = LireEtat("nb_anomalies")
    On Error Resume Next
    ws.Range("B3").Hyperlinks.Delete
    On Error GoTo 0
    If Val(nb) > 0 Then
        s = ChrW(9888) & "  " & LireEtat("txt_anomalies") & "  " & ChrW(8594) & " voir CONTROLES"
        Lien ws.Range("B3"), "'" & F_CTRL & "'!A1", s
        ws.Range("B3").Font.Color = RGB(198, 40, 40)
    Else
        ws.Range("B3").Value = ChrW(10004) & "  Aucune anomalie détectée."
        ws.Range("B3").Font.Color = RGB(46, 125, 50)
    End If
    ws.Range("B3").Font.Bold = True
    ws.Range("B3").Font.Underline = xlUnderlineStyleNone
End Sub

'------------------------------------------------------------------------------
' Tous : un tableau par agent, un tableau par jour
'------------------------------------------------------------------------------
Private Sub EcrireParAgent(ws As Worksheet, v As Variant, ByVal du As Long, ByVal au As Long)
    Dim agents As New Collection, nAg As Long, noms() As String, ix() As Long, i As Long, j As Long, k As Long, r As Long
    Dim ag As String, d As Long, cle As String, jours As New Collection, ch As New Collection
    Dim nJ() As Long, nP() As Long, nOk() As Long, hP() As Double, hR() As Double, nCh() As Long, cpt() As Boolean
    Dim arr() As Variant, n As Long, tot(1 To 12) As Double, kmAg As Double, nbJ As Long, l1 As Long, nCompte As Long

    ReDim noms(1 To 2000): ReDim nJ(1 To 2000): ReDim nP(1 To 2000): ReDim nOk(1 To 2000)
    ReDim hP(1 To 2000): ReDim hR(1 To 2000): ReDim nCh(1 To 2000): ReDim cpt(1 To 2000): ReDim ix(1 To 2000)
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, D_DATE))
        If d >= du And d <= au Then
            ag = Trim$(txt(v(i, D_AGENT)))
            k = Idx(agents, ag)
            If k = 0 And nAg < 2000 Then
                nAg = nAg + 1: k = nAg: Ajout agents, ag, k
                noms(k) = ag: ix(k) = k
            End If
            If k > 0 Then
                cpt(k) = (txt(v(i, D_COMPTE)) <> "0")
                nP(k) = nP(k) + 1
                If txt(v(i, D_POINTEE)) = "1" Then nOk(k) = nOk(k) + 1
                hP(k) = hP(k) + Nombre0(v(i, D_DUREEP))
                hR(k) = hR(k) + Nombre0(v(i, D_HRET))
                If AjoutUnique(jours, ag & "|" & d) Then nJ(k) = nJ(k) + 1
                cle = txt(v(i, D_CLE))
                If cle <> "" Then If AjoutUnique(ch, ag & "|" & cle) Then nCh(k) = nCh(k) + 1
            End If
        End If
    Next i

    r = Z_L1
    ws.Range("B" & r).Value = "PAR AGENT"
    ws.Range("D" & r).Value = "Cliquer sur un nom pour ouvrir sa fiche (prestations dans l'ordre des clics, abonnements)."
    FormatTitre ws.Range("B" & r & ":M" & r), BLEU
    ws.Range("D" & r).Font.Bold = False: ws.Range("D" & r).Font.Size = 9
    r = r + 1
    ws.Range("B" & r & ":M" & r).Value = Array("Agent", "Jours", "Prestations", "% pointées", "Non pointées", "H prévues", _
        "H pointées", "Chantiers", "CA mensuel (" & JoursReference() & " j)", "Km pointés", "Km / jour", "")
    FormatEntete ws.Range("B" & r & ":L" & r)
    ws.Rows(r).RowHeight = 30
    If nAg = 0 Then ws.Range("B" & r + 1).Value = "Aucune prestation sur la période.": Exit Sub
    TriTexte noms, ix, 1, nAg

    ' lignes : agents comptés, total, puis agents hors indicateurs (grisés)
    ReDim arr(1 To nAg + 2, 1 To 11)
    Dim typ() As Long                       ' 1 agent compté, 2 total, 3 libellé, 4 agent non compté
    ReDim typ(1 To nAg + 2)
    l1 = r + 1
    For j = 0 To 1
        If j = 1 And n > 0 Then
            n = n + 1: typ(n) = 2: nCompte = n - 1
            arr(n, 1) = "Total (" & nCompte & " agents)": arr(n, 2) = "": arr(n, 3) = tot(3)
            arr(n, 4) = Pct(tot(4), tot(3)): arr(n, 5) = tot(3) - tot(4): arr(n, 6) = Round(tot(6), 1)
            arr(n, 7) = Round(tot(7), 1): arr(n, 8) = "": arr(n, 9) = Round(tot(9), 0): arr(n, 10) = Round(tot(10), 0)
            arr(n, 11) = ""
        End If
        For i = 1 To nAg
            k = ix(i)
            If (j = 0 And cpt(k)) Or (j = 1 And Not cpt(k)) Then
                If j = 1 Then
                    If n = 0 Then
                        n = n + 1: typ(n) = 3: arr(n, 1) = "Hors indicateurs (réglage dans PARAMETRES) :"
                    ElseIf typ(n) = 2 Then
                        n = n + 1: typ(n) = 3: arr(n, 1) = "Hors indicateurs (réglage dans PARAMETRES) :"
                    End If
                End If
                n = n + 1: typ(n) = IIf(j = 0, 1, 4)
                kmAg = KmPeriode(noms(i), du, au, nbJ, False)
                arr(n, 1) = noms(i): arr(n, 2) = nJ(k): arr(n, 3) = nP(k): arr(n, 4) = Pct(nOk(k), nP(k))
                arr(n, 5) = nP(k) - nOk(k): arr(n, 6) = Round(hP(k), 2): arr(n, 7) = Round(hR(k), 2)
                arr(n, 8) = nCh(k): arr(n, 9) = Round(CAAgent(noms(i)), 0)
                arr(n, 10) = Round(kmAg, 1)
                If nbJ > 0 Then arr(n, 11) = Round(kmAg / nbJ, 1) Else arr(n, 11) = ""
                If j = 0 Then
                    tot(2) = tot(2) + nJ(k): tot(3) = tot(3) + nP(k): tot(4) = tot(4) + nOk(k)
                    tot(6) = tot(6) + hP(k): tot(7) = tot(7) + hR(k)
                    tot(9) = tot(9) + arr(n, 9): tot(10) = tot(10) + kmAg
                End If
            End If
        Next i
    Next j
    ws.Range(ws.Cells(l1, 2), ws.Cells(l1 + n - 1, 12)).Value = arr
    FormatColonnesAgent ws, l1, l1 + n - 1
    FormatLignes ws.Range(ws.Cells(l1, 2), ws.Cells(l1 + n - 1, 12))
    For i = 1 To n
        Select Case typ(i)
            Case 1, 4
                Cliquable ws.Cells(l1 + i - 1, 2), CStr(arr(i, 1))
                ws.Cells(l1 + i - 1, 2).Font.Underline = xlUnderlineStyleNone
                If typ(i) = 1 Then
                    ws.Cells(l1 + i - 1, 2).Font.Color = BLEU
                    ColorerTaux ws.Cells(l1 + i - 1, 5)
                Else
                    ws.Range(ws.Cells(l1 + i - 1, 2), ws.Cells(l1 + i - 1, 12)).Font.Color = RGB(150, 150, 150)
                End If
            Case 2
                With ws.Range(ws.Cells(l1 + i - 1, 2), ws.Cells(l1 + i - 1, 12))
                    .Font.Bold = True
                    .Interior.Color = RGB(244, 246, 249)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
            Case 3
                FormatAide ws.Cells(l1 + i - 1, 2)
                ws.Rows(l1 + i - 1).RowHeight = 24
                ws.Cells(l1 + i - 1, 2).VerticalAlignment = xlBottom
        End Select
    Next i
End Sub

Private Sub FormatColonnesAgent(ws As Worksheet, ByVal l1 As Long, ByVal l2 As Long)
    ws.Range(ws.Cells(l1, 3), ws.Cells(l2, 4)).NumberFormat = "0"
    ws.Range(ws.Cells(l1, 5), ws.Cells(l2, 5)).NumberFormat = "0%"
    ws.Range(ws.Cells(l1, 6), ws.Cells(l2, 6)).NumberFormat = "0"
    ws.Range(ws.Cells(l1, 7), ws.Cells(l2, 8)).NumberFormat = "0.0"
    ws.Range(ws.Cells(l1, 9), ws.Cells(l2, 9)).NumberFormat = "0"
    ws.Range(ws.Cells(l1, 10), ws.Cells(l2, 10)).NumberFormat = "#,##0 " & ChrW(8364)
    ws.Range(ws.Cells(l1, 11), ws.Cells(l2, 12)).NumberFormat = "0"
    ws.Range(ws.Cells(l1, 3), ws.Cells(l2, 12)).HorizontalAlignment = xlRight
End Sub

Private Sub ColorerTaux(c As Range)
    If txt(c.Value) = "" Then Exit Sub
    If c.Value < 0.5 Then
        c.Font.Color = RGB(198, 40, 40)
    ElseIf c.Value < 0.8 Then
        c.Font.Color = RGB(230, 120, 0)
    Else
        c.Font.Color = RGB(46, 125, 50)
    End If
    c.Font.Bold = True
End Sub

Private Sub EcrireParJour(ws As Worksheet, v As Variant, ByVal du As Long, ByVal au As Long)
    Dim jours As New Collection, nJ As Long, dj() As Long, nP() As Long, nOk() As Long, hP() As Double, hR() As Double
    Dim nAg() As Long, vus As New Collection, i As Long, k As Long, d As Long, r As Long, arr() As Variant, ag As String
    ReDim dj(1 To 400): ReDim nP(1 To 400): ReDim nOk(1 To 400): ReDim hP(1 To 400): ReDim hR(1 To 400): ReDim nAg(1 To 400)
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, D_DATE))
        If d >= du And d <= au And txt(v(i, D_COMPTE)) <> "0" Then
            k = Idx(jours, CStr(d))
            If k = 0 And nJ < 400 Then nJ = nJ + 1: k = nJ: Ajout jours, CStr(d), k: dj(k) = d
            If k > 0 Then
                nP(k) = nP(k) + 1
                If txt(v(i, D_POINTEE)) = "1" Then nOk(k) = nOk(k) + 1
                hP(k) = hP(k) + Nombre0(v(i, D_DUREEP)): hR(k) = hR(k) + Nombre0(v(i, D_HRET))
                ag = Trim$(txt(v(i, D_AGENT)))
                If AjoutUnique(vus, d & "|" & ag) Then nAg(k) = nAg(k) + 1
            End If
        End If
    Next i
    r = Z_L1
    ws.Range("O" & r).Value = "PAR JOUR"
    FormatTitre ws.Range("O" & r & ":U" & r), BLEU
    r = r + 1
    ws.Range("O" & r & ":U" & r).Value = Array("Jour", "Prestations", "Pointées", "% pointées", "H prévues", "H pointées", "Agents")
    FormatEntete ws.Range("O" & r & ":U" & r)
    If nJ = 0 Then Exit Sub
    ReDim arr(1 To nJ, 1 To 7)
    For i = 1 To nJ
        k = nJ - i + 1           ' du plus récent au plus ancien (DATA est trié par date)
        arr(i, 1) = CDate(dj(k)): arr(i, 2) = nP(k): arr(i, 3) = nOk(k): arr(i, 4) = Pct(nOk(k), nP(k))
        arr(i, 5) = Round(hP(k), 1): arr(i, 6) = Round(hR(k), 1): arr(i, 7) = nAg(k)
    Next i
    ws.Range(ws.Cells(r + 1, 15), ws.Cells(r + nJ, 21)).Value = arr
    ws.Range(ws.Cells(r + 1, 15), ws.Cells(r + nJ, 15)).NumberFormat = "ddd dd/mm"
    ws.Range(ws.Cells(r + 1, 18), ws.Cells(r + nJ, 18)).NumberFormat = "0%"
    ws.Range(ws.Cells(r + 1, 19), ws.Cells(r + nJ, 20)).NumberFormat = "0"
    ws.Range(ws.Cells(r + 1, 15), ws.Cells(r + nJ, 15)).HorizontalAlignment = xlLeft
    FormatLignes ws.Range(ws.Cells(r + 1, 15), ws.Cells(r + nJ, 21))
    For i = 1 To nJ
        ColorerTaux ws.Cells(r + i, 18)
    Next i
End Sub

'------------------------------------------------------------------------------
' Fiche d'un agent
'------------------------------------------------------------------------------
Private Sub EcrireFicheAgent(ws As Worksheet, v As Variant, ByVal agent As String, ByVal du As Long, ByVal au As Long, _
                             hCle As Collection)
    Dim i As Long, j As Long, k As Long, r As Long, d As Long, n As Long, cle As String, a As Long
    Dim cles As New Collection, nC As Long, cCle() As String, cNom() As String, cPass() As Long, cOk() As Long
    Dim cH() As Double, cDern() As Long, ord() As Long, vals() As Double, arr() As Variant, part As Double, totCA As Double
    Dim lig() As Long, nL As Long, tri() As String, ixL() As Long, l1 As Long, jour As Long, rang As Long
    Dim nJour As Long, okJour As Long, premier As String, dernier As String, kmJ As Double, nbJ As Long, debut As Long

    ' 1) ses chantiers sur la période
    ReDim cCle(1 To 2000): ReDim cNom(1 To 2000): ReDim cPass(1 To 2000): ReDim cOk(1 To 2000)
    ReDim cH(1 To 2000): ReDim cDern(1 To 2000)
    ReDim lig(1 To UBound(v, 1)): ReDim tri(1 To UBound(v, 1)): ReDim ixL(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, D_DATE))
        If d >= du And d <= au And Trim$(txt(v(i, D_AGENT))) = agent Then
            cle = txt(v(i, D_CLE))
            k = Idx(cles, cle)
            If k = 0 And nC < 2000 Then
                nC = nC + 1: k = nC: Ajout cles, cle, k
                cCle(k) = cle: cNom(k) = txt(v(i, D_CHANTIER))
            End If
            If k > 0 Then
                cPass(k) = cPass(k) + 1
                If txt(v(i, D_POINTEE)) = "1" Then cOk(k) = cOk(k) + 1
                cH(k) = cH(k) + Nombre0(v(i, D_DUREEP))
                If d > cDern(k) Then cDern(k) = d
            End If
            ' clé de tri : jour | pointée d'abord, par heure de clic | sinon par heure prévue
            nL = nL + 1: lig(nL) = i: ixL(nL) = nL
            If txt(v(i, D_POINTEE)) = "1" Or Not IsEmpty(v(i, D_DEBR)) Then
                tri(nL) = Format(d, "00000") & "|0|" & Format(Minutes(v(i, D_DEBR)), "0000") & "|" & Format(i, "000000")
            Else
                tri(nL) = Format(d, "00000") & "|1|" & Format(Minutes(v(i, D_DEBP)), "0000") & "|" & Format(i, "000000")
            End If
        End If
    Next i

    r = Z_L1
    ws.Range("B" & r).Value = "FICHE : " & agent
    FormatTitre ws.Range("B" & r & ":L" & r), BLEU
    Cliquable ws.Range("K" & r), ChrW(8592) & " Tous les agents"
    With ws.Range("K" & r)
        .Font.Color = RGB(255, 255, 255): .Font.Bold = True: .Font.Underline = xlUnderlineStyleNone
    End With
    If nC = 0 Then ws.Range("B" & r + 1).Value = "Aucune prestation sur la période.": Exit Sub

    ' 2) ses abonnements / chantiers, triés par CA attribué
    r = r + 2
    ws.Range("B" & r).Value = "SES CHANTIERS ET ABONNEMENTS"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Color = BLEU
    ws.Range("E" & r).Value = "Sa part = ses heures prévues / heures prévues du chantier, sur les " & JoursReference() & " derniers jours."
    FormatAide ws.Range("E" & r)
    r = r + 1
    ws.Range("B" & r & ":L" & r).Value = Array("Chantier", "N° abonnement", "Client facturation", "Fréquence", _
        "CA mensuel chantier", "Sa part", "CA attribué", "Passages", "Pointés", "H prévues", "Dernier passage")
    FormatEntete ws.Range("B" & r & ":L" & r)
    ws.Rows(r).RowHeight = 30
    ReDim ord(1 To nC): ReDim vals(1 To nC): ReDim arr(1 To nC, 1 To 11)
    For k = 1 To nC
        a = Idx(aCle, cCle(k))
        part = 0
        If Nombre0(Valeur(hCle, cCle(k))) > 0 And Idx(colPaire, agent & "|" & cCle(k)) > 0 Then
            part = pH(Idx(colPaire, agent & "|" & cCle(k))) / Nombre0(Valeur(hCle, cCle(k)))
        End If
        If a > 0 Then
            arr(k, 1) = aNom(a): arr(k, 2) = aNum(a): arr(k, 3) = aCli(a): arr(k, 4) = aFreq(a)
            arr(k, 5) = Round(aMens(a), 2): arr(k, 6) = part: arr(k, 7) = Round(aMens(a) * part, 2)
            totCA = totCA + aMens(a) * part
            vals(k) = aMens(a) * part + 0.000001 * cPass(k)
        Else
            arr(k, 1) = cNom(k): arr(k, 2) = "Sans abonnement": arr(k, 3) = "": arr(k, 4) = ""
            arr(k, 5) = "": arr(k, 6) = "": arr(k, 7) = ""
            vals(k) = -1 + 0.000001 * cPass(k)
        End If
        arr(k, 8) = cPass(k): arr(k, 9) = cOk(k): arr(k, 10) = Round(cH(k), 2): arr(k, 11) = CDate(cDern(k))
        ord(k) = k
    Next k
    TriDesc vals, ord, 1, nC
    l1 = r + 1
    For i = 1 To nC
        k = ord(i)
        For j = 1 To 11
            ws.Cells(l1 + i - 1, 1 + j).Value = arr(k, j)
        Next j
        If arr(k, 2) = "Sans abonnement" Then ws.Cells(l1 + i - 1, 3).Font.Color = RGB(198, 40, 40)
    Next i
    ws.Range(ws.Cells(l1, 6), ws.Cells(l1 + nC - 1, 6)).NumberFormat = "#,##0 " & ChrW(8364)
    ws.Range(ws.Cells(l1, 7), ws.Cells(l1 + nC - 1, 7)).NumberFormat = "0%"
    ws.Range(ws.Cells(l1, 8), ws.Cells(l1 + nC - 1, 8)).NumberFormat = "#,##0 " & ChrW(8364)
    ws.Range(ws.Cells(l1, 11), ws.Cells(l1 + nC - 1, 11)).NumberFormat = "0.0"
    ws.Range(ws.Cells(l1, 12), ws.Cells(l1 + nC - 1, 12)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(l1, 3), ws.Cells(l1 + nC - 1, 3)).HorizontalAlignment = xlLeft
    FormatLignes ws.Range(ws.Cells(l1, 2), ws.Cells(l1 + nC - 1, 12))
    r = l1 + nC
    ws.Range("B" & r).Value = "Total"
    ws.Range("H" & r).Value = Round(totCA, 0)
    ws.Range("H" & r).NumberFormat = "#,##0 " & ChrW(8364)
    ws.Range("B" & r & ":L" & r).Font.Bold = True
    ws.Range("B" & r & ":L" & r).Borders(xlEdgeTop).LineStyle = xlContinuous

    ' 3) ses prestations, jour par jour, dans l'ordre des clics
    r = r + 3
    ws.Range("B" & r).Value = "SES PRESTATIONS, DANS L'ORDRE DES CLICS"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Color = BLEU
    ws.Range("F" & r).Value = "Non pointées en fin de journée, à l'heure prévue. Conteneurs : 1er clic = sortie, 2e = rentrée."
    FormatAide ws.Range("F" & r)
    r = r + 1
    ws.Range("B" & r & ":L" & r).Value = Array("Prestation", "Ordre", "Clic début", "Clic fin", "Durée pointée (h)", _
        "Prévu", "Durée prévue (h)", "Type", "N° abonnement", "Fréquence", "Statut")
    FormatEntete ws.Range("B" & r & ":L" & r)
    ws.Rows(r).RowHeight = 30
    TriTexte tri, ixL, 1, nL
    ReDim arr(1 To nL + 400, 1 To 11)
    Dim entetes() As Long, nE As Long
    ReDim entetes(1 To 400)
    l1 = r + 1
    n = 0: jour = 0
    For i = 1 To nL
        j = lig(ixL(i))
        d = SerialDate(v(j, D_DATE))
        If d <> jour Then
            ' en-tête du jour
            jour = d: rang = 0
            ResumeJour v, lig, ixL, i, nL, d, nJour, okJour, premier, dernier
            kmJ = KmPeriode(agent, d, d, nbJ, False)
            n = n + 1: nE = nE + 1: entetes(nE) = n
            arr(n, 1) = StrConv(Format(CDate(d), "dddd dd/mm/yyyy"), vbProperCase)
            arr(n, 2) = nJour & " prestations " & ChrW(183) & " " & okJour & " pointées" & _
                IIf(premier <> "", " " & ChrW(183) & " 1er clic " & premier & " " & ChrW(183) & " dernier " & dernier, "") & _
                IIf(kmJ > 0, " " & ChrW(183) & " " & Format(kmJ, "0.0") & " km pointés", "")
        End If
        n = n + 1
        arr(n, 1) = txt(v(j, D_LIBELLE))
        If arr(n, 1) = "" Then arr(n, 1) = txt(v(j, D_CHANTIER))
        If txt(v(j, D_POINTEE)) = "1" Or Not IsEmpty(v(j, D_DEBR)) Then rang = rang + 1: arr(n, 2) = rang Else arr(n, 2) = "-"
        arr(n, 3) = HHMM(v(j, D_DEBR)): arr(n, 4) = HHMM(v(j, D_FINR))
        If txt(v(j, D_POINTEE)) = "1" Then arr(n, 5) = Nombre0(v(j, D_DUREER)) Else arr(n, 5) = ""
        arr(n, 6) = HHMM(v(j, D_DEBP)) & " - " & HHMM(v(j, D_FINP))
        arr(n, 7) = Nombre0(v(j, D_DUREEP))
        arr(n, 8) = txt(v(j, D_TYPE))
        cle = txt(v(j, D_CLE))
        a = Idx(aCle, cle)
        If a > 0 Then arr(n, 9) = aNum(a): arr(n, 10) = aFreq(a) Else arr(n, 9) = "Sans abonnement": arr(n, 10) = ""
        arr(n, 11) = txt(v(j, D_STATUT))
        If txt(v(j, D_TYPE)) = CONTENEURS And txt(v(j, D_POINTEE)) = "1" Then arr(n, 11) = "OK (sortie / rentrée)"
    Next i
    ws.Range(ws.Cells(l1, 2), ws.Cells(l1 + n - 1, 5)).NumberFormat = "@"
    ws.Range(ws.Cells(l1, 7), ws.Cells(l1 + n - 1, 7)).NumberFormat = "@"
    ws.Range(ws.Cells(l1, 10), ws.Cells(l1 + n - 1, 10)).NumberFormat = "@"
    ws.Range(ws.Cells(l1, 2), ws.Cells(l1 + n - 1, 12)).Value = arr
    ws.Range(ws.Cells(l1, 6), ws.Cells(l1 + n - 1, 6)).NumberFormat = "0.00"
    ws.Range(ws.Cells(l1, 8), ws.Cells(l1 + n - 1, 8)).NumberFormat = "0.00"
    ws.Range(ws.Cells(l1, 3), ws.Cells(l1 + n - 1, 12)).HorizontalAlignment = xlCenter
    FormatLignes ws.Range(ws.Cells(l1, 2), ws.Cells(l1 + n - 1, 12))
    ' couleurs : statut, sans abonnement
    For i = 1 To n
        Select Case Left$(txt(ws.Cells(l1 + i - 1, 12).Value), 2)
            Case "OK": ws.Cells(l1 + i - 1, 12).Font.Color = RGB(46, 125, 50)
            Case "":
            Case Else: ws.Cells(l1 + i - 1, 12).Font.Color = RGB(198, 40, 40)
        End Select
        If txt(ws.Cells(l1 + i - 1, 10).Value) = "Sans abonnement" Then ws.Cells(l1 + i - 1, 10).Font.Color = RGB(198, 40, 40)
    Next i
    For i = 1 To nE
        With ws.Range(ws.Cells(l1 + entetes(i) - 1, 2), ws.Cells(l1 + entetes(i) - 1, 12))
            .Interior.Color = RGB(231, 236, 242)
            .Font.Bold = True
            .Font.Color = BLEU
            .HorizontalAlignment = xlLeft
        End With
        ws.Cells(l1 + entetes(i) - 1, 3).Font.Bold = False
    Next i
End Sub

' Résumé d'une journée de l'agent (lignes triées : i = 1re ligne du jour)
Private Sub ResumeJour(v As Variant, lig() As Long, ixL() As Long, ByVal i As Long, ByVal nL As Long, ByVal d As Long, _
                       nJour As Long, okJour As Long, premier As String, dernier As String)
    Dim k As Long, j As Long, mMin As Long, mMax As Long, m As Long
    nJour = 0: okJour = 0: premier = "": dernier = "": mMin = 9999: mMax = -1
    For k = i To nL
        j = lig(ixL(k))
        If SerialDate(v(j, D_DATE)) <> d Then Exit For
        nJour = nJour + 1
        If txt(v(j, D_POINTEE)) = "1" Then okJour = okJour + 1
        If Not IsEmpty(v(j, D_DEBR)) Then
            m = Minutes(v(j, D_DEBR)): If m < mMin Then mMin = m
            If m > mMax Then mMax = m
        End If
        If Not IsEmpty(v(j, D_FINR)) Then
            m = Minutes(v(j, D_FINR)): If m > mMax Then mMax = m
        End If
    Next k
    If mMax >= 0 Then
        premier = Format(mMin \ 60, "00") & ":" & Format(mMin Mod 60, "00")
        dernier = Format(mMax \ 60, "00") & ":" & Format(mMax Mod 60, "00")
    End If
End Sub

'------------------------------------------------------------------------------
' Abonnements (ABOS) regroupés par chantier
'------------------------------------------------------------------------------
Private Sub ChargerAbos()
    Dim ws As Worksheet, lastR As Long, v As Variant, i As Long, k As Long, cle As String
    Set ws = ThisWorkbook.Worksheets(F_ABOS)
    Set aCle = New Collection: nA = 0
    lastR = DerniereLigne(ws)
    ReDim aNom(1 To Application.WorksheetFunction.Max(1, lastR)): ReDim aNum(1 To UBound(aNom)): ReDim aCli(1 To UBound(aNom))
    ReDim aFreq(1 To UBound(aNom)): ReDim aMens(1 To UBound(aNom)): ReDim aK(1 To UBound(aNom))
    If lastR < 2 Then Exit Sub
    v = ws.Range("A2:H" & lastR).Value
    For i = 1 To UBound(v, 1)
        cle = txt(v(i, 8))
        If cle <> "" Then
            k = Idx(aCle, cle)
            If k = 0 Then
                nA = nA + 1: k = nA: Ajout aCle, cle, k: aK(k) = cle
                aNom(k) = txt(v(i, 2)): aNum(k) = txt(v(i, 1)): aCli(k) = txt(v(i, 3)): aFreq(k) = txt(v(i, 4))
            Else
                aNum(k) = aNum(k) & " + " & txt(v(i, 1))
                If InStr(aFreq(k), txt(v(i, 4))) = 0 Then aFreq(k) = aFreq(k) & " + " & txt(v(i, 4))
            End If
            aMens(k) = aMens(k) + Nombre0(v(i, 6))
        End If
    Next i
End Sub

Private Function CompterAbos(chantiers As Collection) As Long
    ' nombre de chantiers de la collection qui ont un abonnement : on reparcourt les clés connues
    Dim k As Long, n As Long
    For k = 1 To nA
        If Existe(chantiers, CleAbo(k)) Then n = n + 1
    Next k
    CompterAbos = n
End Function

Private Function CleAbo(ByVal k As Long) As String
    CleAbo = aK(k)
End Function

' CA mensuel attribué à l'agent : somme des CA chantier x sa part des heures prévues (période de contrôle)
Private Function CAAgent(ByVal agent As String) As Double
    CAAgent = Nombre0(Valeur(caParAgent, agent))
End Function

' CA mensuel de chaque agent, calculé une fois par actualisation
Private Sub CalculerCAAgents(hCle As Collection)
    Dim k As Long, a As Long, ht As Double
    Set caParAgent = New Collection
    For k = 1 To nPaire
        a = Idx(aCle, pCle(k))
        If a > 0 Then
            ht = Nombre0(Valeur(hCle, pCle(k)))
            If ht > 0 Then AjouterNb caParAgent, pAg(k), aMens(a) * pH(k) / ht
        End If
    Next k
End Sub

Private Sub AjouterPaire(ByVal ag As String, ByVal cle As String, ByVal h As Double)
    Dim k As Long
    k = Idx(colPaire, ag & "|" & cle)
    If k = 0 Then
        nPaire = nPaire + 1: k = nPaire
        If nPaire > UBound(pAg) Then
            ReDim Preserve pAg(1 To 2 * nPaire): ReDim Preserve pCle(1 To 2 * nPaire): ReDim Preserve pH(1 To 2 * nPaire)
        End If
        Ajout colPaire, ag & "|" & cle, k
        pAg(k) = ag: pCle(k) = cle
    End If
    pH(k) = pH(k) + h
End Sub

Private Sub ChargerTournees()
    Dim ws As Worksheet, lastR As Long
    Set ws = ThisWorkbook.Worksheets(F_TOURNEES)
    lastR = DerniereLigne(ws)
    nT = 0: vT = Empty
    If lastR >= 5 Then vT = ws.Range("A5:H" & lastR).Value: nT = UBound(vT, 1)
End Sub

Private Sub AjouterNb(c As Collection, ByVal cle As String, ByVal x As Double)
    Dim y As Double
    If Existe(c, cle) Then
        y = c.Item("k" & cle)
        c.Remove "k" & cle
    End If
    c.Add y + x, "k" & cle
End Sub

' Km pointés (TOURNEES) d'un agent ("" = tous les agents comptés) ; nbJ = nombre de tournées
Public Function KmPeriode(ByVal agent As String, ByVal du As Long, ByVal au As Long, ByRef nbJ As Long, _
                         Optional ByVal seulementComptes As Boolean = False) As Double
    Dim i As Long, d As Long
    nbJ = 0
    For i = 1 To nT
        d = SerialDate(vT(i, 1))
        If d >= du And d <= au Then
            If (agent = "" And (Not seulementComptes Or txt(vT(i, 8)) <> "0")) Or Trim$(txt(vT(i, 2))) = agent Then
                If txt(vT(i, 7)) = "OK" And Nombre0(vT(i, 5)) > 0 Then
                    KmPeriode = KmPeriode + Nombre0(vT(i, 6))
                    nbJ = nbJ + 1
                End If
            End If
        End If
    Next i
End Function
