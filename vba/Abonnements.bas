'==============================================================================
' Abonnements : lecture de l'onglet ABONNEMENTS (export collé tel quel),
' rattachement des prestations (DATA T:V) et table ABOS.
'
' CA mensuel d'un abonnement : MOIS = montant, 4X ANS = montant / 3, 2X ANS = montant / 6...
' Rattachement prestation -> abonnement, dans l'ordre :
'   1. manuel (PARAMETRES, colonnes L:M)   2. n° de projet
'   3. même chantier (nom identique, à moins de 500 m) qu'une prestation déjà rattachée
'   4. nom du chantier = Client Intervention ou N.Ref. d'un seul abonnement
'==============================================================================
Option Explicit

Private Const RAYON_CHANTIER As Double = 0.5

' Chantiers facturés (abonnements actifs regroupés par n° de projet)
Private kN() As Double, kNom() As String, kProj() As String, kMens() As Double, nK As Long
Private kNb() As Long, kDern() As Long, kNbRef() As Long, kHRef() As Double
' Lignes d'abonnement actives
Private rN() As Double, rNom() As String, rCli() As String, rFreq() As String, rProj() As String, rRef() As String
Private rMnt() As Double, rMens() As Double, rK() As Long, rAdr() As String, rDeb() As Variant, rFin() As Variant, nR As Long
Private colK As Collection       ' n° d'abonnement -> indice du chantier
Private colProj As Collection    ' projet -> indice
Private colNom As Collection     ' nom normalisé -> indice (0 si plusieurs chantiers)

Public Function JoursReference() As Long
    JoursReference = CLng(Nombre0(ThisWorkbook.Worksheets(F_PARAM).Range(P_JOURS_REF).Value))
    If JoursReference < 1 Then JoursReference = 28
End Function

Public Function TraiterAbonnements() As String
    Dim dMax As Long, refDeb As Long, msg As String, nAct As Long, totMens As Double
    dMax = DernierJour()
    If dMax = 0 Then dMax = CLng(Date)
    refDeb = dMax - JoursReference() + 1
    nK = 0: nR = 0
    Set colK = New Collection: Set colProj = New Collection: Set colNom = New Collection
    msg = LireAbos(ThisWorkbook.Worksheets(F_ABO), refDeb, dMax, nAct, totMens)
    If msg = "" Then msg = "Abonnements : " & nAct & " actifs, " & Euros(totMens) & " HT / mois."
    EcrireEtat "nb_abos", CStr(nAct)
    msg = msg & vbCrLf & Rattacher(refDeb, dMax)
    EcrireAbos
    TraiterAbonnements = msg
End Function

' Lit l'onglet ABONNEMENTS. Renvoie "" si OK, sinon un message d'erreur.
Private Function LireAbos(ws As Worksheet, ByVal refDeb As Long, ByVal dMax As Long, nAct As Long, totMens As Double) As String
    Dim v As Variant, lastR As Long, lastC As Long, i As Long, j As Long, h As Long, e As String
    Dim cN As Long, cDeb As Long, cFin As Long, cProj As Long, cFreq As Long, cCli As Long, cCI As Long, cRef As Long
    Dim cMnt As Long, cAdr As Long, cCP As Long, cLoc As Long
    Dim n As Variant, actif As Boolean, mnt As Variant, proj As String, nom As String, k As Long, mens As Double, fq As String

    lastR = DerniereLigne(ws): lastC = DerniereColonne(ws)
    If lastR < 2 Then LireAbos = "Abonnements : l'onglet ABONNEMENTS est vide (y coller l'export des abonnements).": GoTo Vide
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    For i = 1 To IIf(lastR < 10, lastR, 10)
        For j = 1 To lastC
            If LCase$(Trim$(txt(v(i, j)))) = "n" & ChrW(176) Then h = i: Exit For
        Next j
        If h > 0 Then Exit For
    Next i
    If h = 0 Then LireAbos = "Abonnements (erreur) : colonne « N" & ChrW(176) & " » introuvable dans ABONNEMENTS.": GoTo Vide
    For j = 1 To lastC
        e = LCase$(Trim$(txt(v(h, j))))
        Select Case e
            Case "n" & ChrW(176): cN = j
            Case "date d" & ChrW(233) & "but": cDeb = j
            Case "date fin": cFin = j
            Case "projet": cProj = j
            Case "fr" & ChrW(233) & "quence": cFreq = j
            Case "client facturation": cCli = j
            Case "client intervention": cCI = j
            Case "n.ref.", "n.ref", "n. ref.": cRef = j
            Case "mnt htx", "mnt ht": cMnt = j
            Case "adresse intervention": cAdr = j
            Case "cp": cCP = j
            Case "loc. intervention": cLoc = j
        End Select
    Next j
    If cMnt = 0 Then LireAbos = "Abonnements (erreur) : colonne « Mnt Htx » introuvable dans ABONNEMENTS.": GoTo Vide

    ReDim kN(1 To lastR): ReDim kNom(1 To lastR): ReDim kProj(1 To lastR): ReDim kMens(1 To lastR)
    ReDim kNb(1 To lastR): ReDim kDern(1 To lastR): ReDim kNbRef(1 To lastR): ReDim kHRef(1 To lastR)
    ReDim rN(1 To lastR): ReDim rNom(1 To lastR): ReDim rCli(1 To lastR): ReDim rFreq(1 To lastR): ReDim rProj(1 To lastR)
    ReDim rRef(1 To lastR): ReDim rMnt(1 To lastR): ReDim rMens(1 To lastR): ReDim rK(1 To lastR): ReDim rAdr(1 To lastR)
    ReDim rDeb(1 To lastR): ReDim rFin(1 To lastR)

    For i = h + 1 To lastR
        actif = False
        n = Nombre(v(i, cN))
        If Not IsEmpty(n) And Trim$(txt(v(i, cN))) Like "*[0-9]*" Then
            If n > 0 Then
                actif = True
                If cFin > 0 Then
                    If EstUneDate(v(i, cFin)) Then
                        If SerialDate(v(i, cFin)) < refDeb Then actif = False      ' terminé
                    End If
                End If
                If cDeb > 0 Then
                    If EstUneDate(v(i, cDeb)) Then
                        If SerialDate(v(i, cDeb)) > dMax Then actif = False        ' pas encore commencé
                    End If
                End If
            End If
        End If
        If actif Then
            mnt = Nombre(v(i, cMnt)): If IsEmpty(mnt) Then mnt = 0
            fq = "": If cFreq > 0 Then fq = Trim$(txt(v(i, cFreq)))
            mens = CDbl(mnt) * Facteur(fq)
            proj = "": If cProj > 0 Then proj = UCase$(Trim$(txt(v(i, cProj))))
            nom = ""
            If cCI > 0 Then nom = Trim$(txt(v(i, cCI)))
            If nom = "" And cRef > 0 Then nom = Trim$(txt(v(i, cRef)))
            If nom = "" And cCli > 0 Then nom = Trim$(txt(v(i, cCli)))
            nAct = nAct + 1: totMens = totMens + mens

            k = 0
            If proj <> "" Then k = Idx(colProj, proj)
            If k = 0 Then
                nK = nK + 1: k = nK
                kN(k) = n: kNom(k) = nom: kProj(k) = proj
                If proj <> "" Then Ajout colProj, proj, k
            End If
            kMens(k) = kMens(k) + mens
            Ajout colK, CStr(n), k
            nR = nR + 1
            rN(nR) = n: rNom(nR) = nom: rProj(nR) = proj: rMnt(nR) = CDbl(mnt): rMens(nR) = mens: rK(nR) = k: rFreq(nR) = fq
            If cCli > 0 Then rCli(nR) = Trim$(txt(v(i, cCli)))
            If cRef > 0 Then rRef(nR) = Trim$(txt(v(i, cRef)))
            If cAdr > 0 Then rAdr(nR) = Trim$(txt(v(i, cAdr)))
            If cCP > 0 Then rAdr(nR) = Trim$(rAdr(nR) & ", " & txt(v(i, cCP)))
            If cLoc > 0 Then rAdr(nR) = Trim$(rAdr(nR) & " " & txt(v(i, cLoc)))
            If cDeb > 0 Then If EstUneDate(v(i, cDeb)) Then rDeb(nR) = CDate(SerialDate(v(i, cDeb)))
            If cFin > 0 Then If EstUneDate(v(i, cFin)) Then rFin(nR) = CDate(SerialDate(v(i, cFin)))
            If cCI > 0 Then AjoutNom Norm(txt(v(i, cCI))), k
            If cRef > 0 Then AjoutNom Norm(txt(v(i, cRef))), k
        End If
    Next i
    Exit Function
Vide:
    ReDim kN(1 To 1): ReDim kNom(1 To 1): ReDim kProj(1 To 1): ReDim kMens(1 To 1)
    ReDim kNb(1 To 1): ReDim kDern(1 To 1): ReDim kNbRef(1 To 1): ReDim kHRef(1 To 1)
End Function

' Coefficient mensuel selon la fréquence de facturation (MOIS, 4X ANS, 2X ANS, TRIM, SEM, AN)
Public Function Facteur(ByVal f As String) As Double
    Dim u As String, p As Long
    u = UCase$(Replace(Trim$(f), " ", ""))
    Facteur = 1
    If u = "" Or u = "MOIS" Or u = "MENSUEL" Then Exit Function
    p = InStr(u, "XAN")
    If p > 1 Then
        If IsNumeric(Left$(u, p - 1)) Then Facteur = Val(Left$(u, p - 1)) / 12: Exit Function
    End If
    If u Like "TRIM*" Then Facteur = 1 / 3: Exit Function
    If u Like "SEM*" Then Facteur = 1 / 6: Exit Function
    If u = "AN" Or u Like "ANN*" Then Facteur = 1 / 12: Exit Function
End Function

' Nom -> chantier ; 0 si le nom désigne plusieurs chantiers
Private Sub AjoutNom(ByVal cle As String, ByVal k As Long)
    If cle = "" Then Exit Sub
    If Not Existe(colNom, cle) Then
        Ajout colNom, cle, k
    ElseIf Idx(colNom, cle) <> k And Idx(colNom, cle) <> 0 Then
        colNom.Remove "k" & cle
        Ajout colNom, cle, 0
    End If
End Sub

'------------------------------------------------------------------------------
' Rattachement (DATA : T = n° abonnement, U = mode, V = clé chantier)
'------------------------------------------------------------------------------
Private Function Rattacher(ByVal refDeb As Long, ByVal dMax As Long) As String
    Dim wsD As Worksheet, wsP As Worksheet, lastR As Long, v As Variant, i As Long, j As Long, k As Long
    Dim sortie As Variant, man As New Collection, pv As Variant, nomN As String, nbL As Long
    Dim noms As New Collection, lst() As String, nL As Long, c() As String, e() As String, p As Long
    Dim la As Variant, lo As Variant, nServis As Long, caServi As Double, nSans As Long, caSans As Double
    Dim nManKo As Long, nRef As Long, nRefOk As Long, statut() As Variant, kk() As Long, d As Long

    Set wsD = ThisWorkbook.Worksheets(F_DATA)
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)

    ' Rattachements manuels (PARAMETRES L:M, statut en N)
    lastR = DerniereLigne(wsP)
    If lastR >= P_MAN1 Then
        pv = wsP.Range(wsP.Cells(P_MAN1, 12), wsP.Cells(lastR, 13)).Value
        ReDim statut(1 To UBound(pv, 1), 1 To 1)
        For i = 1 To UBound(pv, 1)
            If Trim$(txt(pv(i, 1))) <> "" And Trim$(txt(pv(i, 2))) <> "" Then
                k = Idx(colK, CStr(Nombre(pv(i, 2))))
                If k > 0 Then
                    Ajout man, "p" & UCase$(Trim$(txt(pv(i, 1)))), k
                    Ajout man, "c" & Norm(txt(pv(i, 1))), k
                    statut(i, 1) = "OK"
                Else
                    nManKo = nManKo + 1
                    statut(i, 1) = "N" & ChrW(176) & " inconnu ou abonnement terminé"
                End If
            End If
        Next i
        wsP.Range(wsP.Cells(P_MAN1, 14), wsP.Cells(lastR, 14)).Value = statut
    End If

    lastR = DerniereLigne(wsD)
    If lastR < 2 Then Rattacher = "Aucune prestation intégrée.": Exit Function
    v = wsD.Range(wsD.Cells(2, 1), wsD.Cells(lastR, NB_COLS)).Value
    nbL = UBound(v, 1)
    ReDim sortie(1 To nbL, 1 To 3)
    ReDim kk(1 To nbL)

    ' 1 + 2 : manuel, puis projet
    For i = 1 To nbL
        k = Idx(man, "p" & UCase$(Trim$(txt(v(i, D_PROJET)))))
        If k = 0 Then k = Idx(man, "c" & NormC(txt(v(i, D_CHANTIER))))
        If k > 0 Then
            kk(i) = k: sortie(i, 2) = "Manuel"
        Else
            k = Idx(colProj, UCase$(Trim$(txt(v(i, D_PROJET)))))
            If k > 0 Then kk(i) = k: sortie(i, 2) = "Projet"
        End If
    Next i

    ' 3 : même chantier (nom + < 500 m) qu'une prestation déjà rattachée
    ReDim lst(1 To nbL)
    For i = 1 To nbL
        If kk(i) > 0 Then
            la = Nombre(v(i, D_LAT)): lo = Nombre(v(i, D_LON))
            If Not IsEmpty(la) And Not IsEmpty(lo) Then
                nomN = NormC(txt(v(i, D_CHANTIER)))
                If nomN <> "" Then
                    j = Idx(noms, nomN)
                    If j = 0 Then nL = nL + 1: j = nL: Ajout noms, nomN, j
                    If Len(lst(j)) < 2000 And InStr(lst(j), "|" & kk(i) & "|") = 0 Then lst(j) = lst(j) & Num5(la) & ";" & Num5(lo) & ";" & kk(i) & "|"
                End If
            End If
        End If
    Next i
    For i = 1 To nbL
        If kk(i) = 0 Then
            la = Nombre(v(i, D_LAT)): lo = Nombre(v(i, D_LON))
            If Not IsEmpty(la) And Not IsEmpty(lo) Then
                j = Idx(noms, NormC(txt(v(i, D_CHANTIER))))
                If j > 0 Then
                    c = Split(lst(j), "|")
                    For p = 0 To UBound(c)
                        If c(p) <> "" Then
                            e = Split(c(p), ";")
                            If DistKm(CDbl(la), CDbl(lo), Val(e(0)), Val(e(1))) < RAYON_CHANTIER Then
                                kk(i) = CLng(Val(e(2))): sortie(i, 2) = "Chantier": Exit For
                            End If
                        End If
                    Next p
                End If
            End If
        End If
    Next i

    ' 4 : nom du chantier = Client Intervention ou N.Ref. (un seul chantier)
    For i = 1 To nbL
        If kk(i) = 0 Then
            k = Idx(colNom, NormC(txt(v(i, D_CHANTIER))))
            If k = 0 Then k = Idx(colNom, NormC(txt(v(i, D_LIBELLE))))
            If k > 0 Then kk(i) = k: sortie(i, 2) = "Nom"
        End If
    Next i

    ' Écriture + statistiques par chantier
    For i = 1 To nbL
        d = SerialDate(v(i, D_DATE))
        If kk(i) > 0 Then
            sortie(i, 1) = kN(kk(i))
            sortie(i, 3) = CStr(kN(kk(i)))
            k = kk(i)
            kNb(k) = kNb(k) + 1
            If d > kDern(k) Then kDern(k) = d
            If d >= refDeb And d <= dMax Then
                kNbRef(k) = kNbRef(k) + 1
                kHRef(k) = kHRef(k) + Nombre0(v(i, D_DUREEP))
            End If
        ElseIf Trim$(txt(v(i, D_PROJET))) <> "" Then
            sortie(i, 3) = "P:" & UCase$(Trim$(txt(v(i, D_PROJET))))
        Else
            sortie(i, 3) = "C:" & NormC(txt(v(i, D_CHANTIER)))
        End If
        If d >= refDeb And d <= dMax And Not EstAgentFictif(txt(v(i, D_AGENT))) Then
            nRef = nRef + 1
            If kk(i) > 0 Then nRefOk = nRefOk + 1
        End If
    Next i
    wsD.Range(wsD.Cells(2, D_ABO), wsD.Cells(lastR, D_CLE)).NumberFormat = "@"
    wsD.Range(wsD.Cells(2, D_ABO), wsD.Cells(lastR, D_ABO)).NumberFormat = "0"
    wsD.Range(wsD.Cells(2, D_ABO), wsD.Cells(lastR, D_CLE)).Value = sortie

    For k = 1 To nK
        If kNbRef(k) > 0 Then
            nServis = nServis + 1: caServi = caServi + kMens(k)
        Else
            nSans = nSans + 1: caSans = caSans + kMens(k)
        End If
    Next k
    If nK = 0 Then
        Rattacher = "Rattachement : aucun abonnement actif."
    Else
        Rattacher = "Rattachement : " & Format(nRefOk / IIf(nRef = 0, 1, nRef), "0%") & " des prestations des " & JoursReference() & _
            " derniers jours. " & nSans & " chantier(s) facturé(s) sans prestation (" & Euros(caSans) & " / mois) : voir CONTROLES."
        If nManKo > 0 Then Rattacher = Rattacher & vbCrLf & nManKo & " rattachement(s) manuel(s) vers un n" & ChrW(176) & _
            " d'abonnement inconnu ou terminé (PARAMETRES)."
    End If
End Function

' ABOS : une ligne par abonnement actif (lue par DASHBOARD et CONTROLES)
Private Sub EcrireAbos()
    Dim ws As Worksheet, arr As Variant, i As Long, lastR As Long, k As Long
    Set ws = ThisWorkbook.Worksheets(F_ABOS)
    lastR = DerniereLigne(ws)
    If lastR >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 16)).ClearContents
    ws.Range("A1:P1").Value = Array("N° abonnement", "Chantier", "Client facturation", "Fréquence", "Montant HT", _
        "CA HT mensuel", "Projet", "Clé chantier", "Adresse", "N.Ref.", "Prestations (tout)", "Dernière prestation", _
        "Prestations (période de contrôle)", "Heures prévues (période de contrôle)", "Date début", "Date fin")
    If nR = 0 Then Exit Sub
    ReDim arr(1 To nR, 1 To 16)
    For i = 1 To nR
        k = rK(i)
        arr(i, 1) = rN(i): arr(i, 2) = rNom(i): arr(i, 3) = rCli(i): arr(i, 4) = rFreq(i)
        arr(i, 5) = Round(rMnt(i), 2): arr(i, 6) = Round(rMens(i), 2): arr(i, 7) = rProj(i)
        arr(i, 8) = CStr(kN(k)): arr(i, 9) = rAdr(i): arr(i, 10) = rRef(i)
        arr(i, 11) = kNb(k)
        If kDern(k) > 0 Then arr(i, 12) = CDate(kDern(k))
        arr(i, 13) = kNbRef(k): arr(i, 14) = Round(kHRef(k), 2)
        arr(i, 15) = rDeb(i): arr(i, 16) = rFin(i)
    Next i
    ws.Range(ws.Cells(2, 8), ws.Cells(nR + 1, 8)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 1), ws.Cells(nR + 1, 16)).Value = arr
    ws.Range(ws.Cells(2, 12), ws.Cells(nR + 1, 12)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 15), ws.Cells(nR + 1, 16)).NumberFormat = "dd/mm/yyyy"
End Sub

'------------------------------------------------------------------------------
' Corrections saisies dans CONTROLES (colonne J des prestations sans abonnement)
' -> rattachements manuels de PARAMETRES (L = projet ou chantier, M = n°)
'------------------------------------------------------------------------------
Public Sub RecupererCorrections()
    Dim ws As Worksheet, wsP As Worksheet, lastR As Long, i As Long, r As Long, cle As String, n As Variant
    Set ws = ThisWorkbook.Worksheets(F_CTRL)
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)
    lastR = DerniereLigne(ws)
    For i = 1 To lastR
        If CStr(ws.Cells(i, 1).Value) = "SANSABO" Then
            n = Nombre(ws.Cells(i, 10).Value)
            If Not IsEmpty(n) Then
                cle = Trim$(txt(ws.Cells(i, 2).Value))
                If cle = "" Then cle = Trim$(txt(ws.Cells(i, 3).Value))
                If cle <> "" Then
                    r = P_MAN1
                    Do While Trim$(txt(wsP.Cells(r, 12).Value)) <> "" Or Trim$(txt(wsP.Cells(r, 13).Value)) <> ""
                        If UCase$(Trim$(txt(wsP.Cells(r, 12).Value))) = UCase$(cle) Then Exit Do
                        r = r + 1
                    Loop
                    wsP.Cells(r, 12).NumberFormat = "@"
                    wsP.Cells(r, 12).Value = cle
                    wsP.Cells(r, 13).Value = n
                End If
                ws.Cells(i, 10).ClearContents
            End If
        End If
    Next i
End Sub

' Après une saisie dans CONTROLES : on rattache tout de suite (sans relire les onglets jours)
Public Sub MajRattachements()
    Dim calcAvant As Long
    calcAvant = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Fin
    Application.StatusBar = "Rattachement..."
    RecupererCorrections
    TraiterAbonnements
    EcrireControles ""
    ActualiserDashboard
Fin:
    Application.Calculation = calcAvant
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function Num5(ByVal x As Variant) As String
    Num5 = Replace(CStr(Round(CDbl(x), 5)), ",", ".")
End Function
