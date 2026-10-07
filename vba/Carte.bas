'==============================================================================
' Carte (page HTML Leaflet, onglet masqué MODELE_CARTE) et kilomètres pointés
'
'   OuvrirCarte  : carte de la période et de l'agent choisis sur le DASHBOARD
'   CalculerKm   : km par tournée (un agent, un jour), dans l'ordre des pointages
'                  (conteneurs : 1er clic = sortie, 2e clic = rentrée, donc 2 passages).
'                  Distances routières OpenStreetMap (OSRM), mémorisées dans DISTANCES.
'   Départ : domicile (si « Départ du domicile » = Oui et adresse localisée), sinon agence.
'   Les choix faits sur la carte (agence matin / soir, maison) reviennent dans PARAMETRES.
'==============================================================================
Option Explicit

Private Const MAX_PTS As Long = 80              ' points par requête d'itinéraire
Private Const URL_ROUTE As String = "https://router.project-osrm.org/table/v1/driving/"
Private Const T_LIGNE1 As Long = 5              ' 1re ligne de données de TOURNEES
Private Const T_NBCOLS As Long = 15             ' A:O

Private gCache As Collection                    ' trajet "lat,lon>lat,lon" -> km
Private gAgLat As Double, gAgLon As Double, gAg As String
Private gRayon As Double, gDom As Collection, gMatin As Collection, gSoir As Collection, gRetour As Boolean
Private gHorsLigne As Boolean, gEchecs As Long, gLots As Long, gLotsKo As Long
Private gNouvK() As String, gNouvV() As Double, gNbNouv As Long
Private gManqA() As String, gManqB() As String, gNbManq As Long, gManq As Collection
Private gVeille As Boolean, gProchain As Date, gFinVeille As Date

'------------------------------------------------------------------------------
' Carte
'------------------------------------------------------------------------------
Public Sub OuvrirCarte()
    Dim wb As Workbook, wsD As Worksheet, wsData As Worksheet
    Dim du As Long, au As Long, t As Long, lastR As Long, v As Variant
    Dim lignes() As String, n As Long, i As Long, jour As Long
    Dim jMin As Long, jMax As Long, json As String, html As String, chemin As String, agent As String
    Dim pts As New Collection

    Set wb = ThisWorkbook
    Set wsD = wb.Worksheets(F_DASH)
    Set wsData = wb.Worksheets(F_DATA)
    If Not EstUneDate(wsD.Range(CELL_DU).Value) Or Not EstUneDate(wsD.Range(CELL_AU).Value) Then
        MsgBox "Renseignez les dates Du / Au du dashboard.", vbExclamation: Exit Sub
    End If
    du = SerialDate(wsD.Range(CELL_DU).Value): au = SerialDate(wsD.Range(CELL_AU).Value)
    If du > au Then t = du: du = au: au = t
    agent = txt(wsD.Range(CELL_AGENT).Value)

    lastR = DerniereLigne(wsData)
    If lastR < 2 Then MsgBox "Aucune prestation : collez d'abord un extract de validations.", vbExclamation: Exit Sub
    Application.StatusBar = "Préparation de la carte..."
    v = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastR, 17)).Value

    ReDim lignes(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        jour = SerialDate(v(i, 1))
        If jour > 0 And Nombre0(v(i, 16)) <> 0 And Nombre0(v(i, 17)) <> 0 Then
            n = n + 1
            If jMin = 0 Or jour < jMin Then jMin = jour
            If jour > jMax Then jMax = jour
            AjouterCle pts, Pt(Nombre0(v(i, 16)), Nombre0(v(i, 17)))
            ' [jour, agent, chantier, libellé, type, début prévu, fin prévue, début réel, fin réelle, enregistré, lat, lon]
            lignes(n) = "[" & JS(Iso(jour)) & "," & JS(txt(v(i, 3))) & "," & JS(txt(v(i, 5))) & "," & _
                JS(txt(v(i, 6))) & "," & JS(txt(v(i, 8))) & "," & JS(HHMM(v(i, 9))) & "," & _
                JS(HHMM(v(i, 10))) & "," & JS(HHMM(v(i, 12))) & "," & JS(HHMM(v(i, 13))) & "," & _
                IIf(Nombre0(v(i, 15)) = 1, "1", "0") & "," & Num(v(i, 16)) & "," & Num(v(i, 17)) & "]"
        End If
    Next i
    If n = 0 Then
        Application.StatusBar = False
        MsgBox "Aucune prestation géolocalisée.", vbExclamation
        Exit Sub
    End If
    ReDim Preserve lignes(1 To n)

    If au > jMax Or au < jMin Then au = jMax
    If du < jMin Or du > au Then du = au
    json = "{""du"":" & JS(Iso(du)) & ",""au"":" & JS(Iso(au)) & ",""min"":" & JS(Iso(jMin)) & _
           ",""max"":" & JS(Iso(jMax)) & ",""agent"":" & JS(agent) & _
           ",""rows"":[" & Join(lignes, ",") & "],""km"":" & JsonKm(pts) & "}"

    html = LireModele(wb.Worksheets(F_MODELE))
    If InStr(html, "/*__INIT__*/null") = 0 Then
        Application.StatusBar = False
        MsgBox "Modèle de carte introuvable (onglet masqué MODELE_CARTE).", vbCritical: Exit Sub
    End If
    html = Replace(html, "/*__INIT__*/null", json)

    chemin = DossierTemp() & "carte_prestations.html"
    Dim f As Integer
    f = FreeFile
    Open chemin For Output As #f
    Print #f, html;
    Close #f
    Application.StatusBar = False

    #If Mac Then
        ThisWorkbook.FollowHyperlink "file://" & chemin
    #Else
        Shell "explorer.exe """ & chemin & """", vbNormalFocus
    #End If
    DemarrerVeille
End Sub

' Agence, réglages des agents, km des tournées et distances connues entre les points affichés
Private Function JsonKm(pts As Collection) As String
    Dim wsT As Worksheet, wsDi As Worksheet, wsP As Worksheet, v As Variant, i As Long, lastR As Long
    Dim ag() As String, nAg As Long, tr() As String, nTr As Long, lg() As String, nLg As Long
    Dim la As Variant, lo As Variant, nom As String, c() As String, dep As Boolean, adr As String

    On Error GoTo Echec
    JsonKm = "null"
    If Not LireParametres() Then Exit Function
    gAg = Pt(gAgLat, gAgLon)
    AjouterCle pts, gAg
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)

    ' Agents : "nom":[agence matin, agence soir, départ maison, adresse, lat, lon]
    lastR = DerniereLigne(wsP)
    ReDim ag(1 To Application.WorksheetFunction.Max(1, lastR))
    For i = P_AG1 To lastR
        nom = Trim$(txt(wsP.Cells(i, 2).Value))
        If nom <> "" Then
            dep = OuiNon(wsP.Cells(i, 5).Value, False)
            adr = Trim$(txt(wsP.Cells(i, 4).Value))
            la = Nombre0(wsP.Cells(i, 8).Value): lo = Nombre0(wsP.Cells(i, 9).Value)
            If la <> 0 And lo <> 0 Then AjouterCle pts, Pt(CDbl(la), CDbl(lo))
            nAg = nAg + 1
            ag(nAg) = JS(nom) & ":[" & IIf(OuiNon(wsP.Cells(i, 6).Value, False), "1", "0") & "," & _
                IIf(OuiNon(wsP.Cells(i, 7).Value, False), "1", "0") & "," & IIf(dep, "1", "0") & "," & _
                JS(adr) & "," & Num(la) & "," & Num(lo) & "]"
        End If
    Next i

    ' Tournées : [jour, agent, km planning (non calculé), km pointés, compté, en erreur]
    Set wsT = ThisWorkbook.Worksheets(F_TOURNEES)
    lastR = DerniereLigne(wsT)
    If lastR >= T_LIGNE1 Then
        v = wsT.Range(wsT.Cells(T_LIGNE1, 1), wsT.Cells(lastR, 8)).Value
        ReDim tr(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            If EstUneDate(v(i, 1)) And txt(v(i, 2)) <> "" Then
                nTr = nTr + 1
                tr(nTr) = "[" & JS(Iso(SerialDate(v(i, 1)))) & "," & JS(Trim$(txt(v(i, 2)))) & ",0," & _
                          Num(v(i, 6)) & "," & IIf(txt(v(i, 8)) = "0", "0", "1") & "," & _
                          IIf(txt(v(i, 7)) = "OK", "0", "1") & "]"
            End If
        Next i
    End If

    ' Distances routières connues entre points de la carte
    Set wsDi = ThisWorkbook.Worksheets(F_DIST)
    lastR = DerniereLigne(wsDi)
    If lastR >= 2 Then
        v = wsDi.Range(wsDi.Cells(2, 1), wsDi.Cells(lastR, 2)).Value
        ReDim lg(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            c = Split(txt(v(i, 1)), ">")
            If UBound(c) = 1 Then
                If Existe(pts, c(0)) And Existe(pts, c(1)) And Not IsEmpty(Nombre(v(i, 2))) Then
                    nLg = nLg + 1
                    lg(nLg) = JS(txt(v(i, 1))) & ":" & Num(v(i, 2))
                End If
            End If
        Next i
    End If

    JsonKm = "{""ag"":[" & Num(gAgLat) & "," & Num(gAgLon) & "],""agNom"":" & JS(txt(wsP.Range(P_ADR_AGENCE).Value)) & _
             ",""rayon"":" & Num(gRayon) & ",""retour"":" & IIf(gRetour, "true", "false") & _
             ",""agents"":{" & Joindre(ag, nAg) & "}" & _
             ",""tours"":[" & Joindre(tr, nTr) & "],""legs"":{" & Joindre(lg, nLg) & "}}"
    Exit Function
Echec:
    JsonKm = "null"
End Function

Private Function Joindre(t() As String, ByVal n As Long) As String
    If n = 0 Then Exit Function
    ReDim Preserve t(1 To n)
    Joindre = Join(t, ",")
End Function

Private Sub AjouterCle(c As Collection, ByVal cle As String)
    On Error Resume Next
    c.Add 1, "k" & cle
End Sub

Private Function LireModele(wsM As Worksheet) As String
    Dim i As Long, s As String
    For i = 1 To 500
        If Len(wsM.Cells(i, 1).Value) = 0 Then Exit For
        s = s & wsM.Cells(i, 1).Value
    Next i
    LireModele = s
End Function

Private Function DossierTemp() As String
    Dim d As String
    #If Mac Then
        d = Environ("TMPDIR")
        If d = "" Then d = Environ("HOME") & "/"
        If Right$(d, 1) <> "/" Then d = d & "/"
    #Else
        d = Environ("TEMP")
        If d = "" Then d = Environ("TMP")
        If d = "" Then d = CurDir$()
        If Right$(d, 1) <> "\" Then d = d & "\"
    #End If
    DossierTemp = d
End Function

'------------------------------------------------------------------------------
' Kilomètres pointés
'------------------------------------------------------------------------------
' Bouton de PARAMETRES : localiser les adresses et tout recalculer
Public Sub RecalculerKm()
    Dim msg As String, calcAvant As Long
    calcAvant = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    msg = CalculerKm("*")
    EcrireEtat "sig_km", SignatureKm()
    ActualiserDashboard
    Application.Calculation = calcAvant
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If msg <> "" Then MsgBox msg, IIf(InStr(1, msg, "erreur", vbTextCompare) > 0, vbExclamation, vbInformation)
End Sub

' joursCle = "*" (tous les jours de DATA) ou "|46287|46288|". Renvoie le bilan.
Public Function CalculerKm(ByVal joursCle As String) As String
    Dim wsData As Worksheet, lastR As Long, v As Variant, i As Long, j As Long, n As Long, r As Long
    Dim cles() As String, ix() As Long, jour As Long, agent As String, dom As String
    Dim pointe() As String, nPointe As Long, nPresta As Long, nPoint As Long, nIgn As Long, sortie As Variant, nT As Long
    Dim enErreur As Boolean, k As Long, kmR(1 To 6) As Variant, nErr As Long, adrKo As String, bilan As String

    On Error GoTo Echec
    Set wsData = ThisWorkbook.Worksheets(F_DATA)
    adrKo = GeocoderAdresses()
    If Not LireParametres() Then
        EcrireTournees Empty, 0, "*", ""
        CalculerKm = "Kilomètres non calculés (erreur) : l'adresse de l'agence n'est pas localisée (PARAMETRES)."
        Exit Function
    End If
    ChargerAgents
    lastR = DerniereLigne(wsData)
    If lastR < 2 Then EcrireTournees Empty, 0, "*", "": Exit Function
    v = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastR, NB_COLS)).Value

    ' 1) Lignes à traiter, triées par jour | agent | ordre d'origine
    ReDim cles(1 To UBound(v, 1)): ReDim ix(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        agent = Trim$(txt(v(i, D_AGENT)))
        jour = SerialDate(v(i, D_DATE))
        If agent <> "" And jour > 0 Then
            If joursCle = "*" Or InStr(1, joursCle, "|" & jour & "|") > 0 Then
                n = n + 1
                cles(n) = Format(jour, "000000") & "|" & agent & "|" & Format(i, "000000")
                ix(n) = i
            End If
        End If
    Next i
    If n = 0 Then EcrireTournees Empty, 0, joursCle, JoursPresents(v): Exit Function
    ReDim Preserve cles(1 To n): ReDim Preserve ix(1 To n)
    TriTexte cles, ix, 1, n

    ChargerCache
    gAg = Pt(gAgLat, gAgLon)
    gHorsLigne = False: gEchecs = 0: gLots = 0: gLotsKo = 0: gNbNouv = 0

    ' 2) Trajets inconnus -> itinéraire routier OpenStreetMap, par lots
    Set gManq = New Collection: gNbManq = 0
    i = 1
    Do While i <= n
        j = FinGroupe(cles, i, n)
        Tournee v, ix, i, j, pointe, nPointe, nPresta, nPoint, nIgn
        dom = DomicileDe(Trim$(txt(v(ix(i), D_AGENT))))
        BesoinsSequence pointe, nPointe, dom
        i = j + 1
    Loop
    If gNbManq > 0 Then CalculerManquants
    EcrireCache

    ' 3) Une ligne par tournée ; un trajet sans itinéraire = tournée en erreur
    ReDim sortie(1 To n, 1 To T_NBCOLS)
    i = 1
    Do While i <= n
        j = FinGroupe(cles, i, n)
        Tournee v, ix, i, j, pointe, nPointe, nPresta, nPoint, nIgn
        r = ix(i)
        agent = Trim$(txt(v(r, D_AGENT)))
        dom = DomicileDe(agent)
        enErreur = False
        KmTournee pointe, nPointe, dom, kmR, enErreur
        nT = nT + 1
        sortie(nT, 1) = CDate(SerialDate(v(r, D_DATE)))
        sortie(nT, 2) = agent
        sortie(nT, 3) = DepartDe(agent, dom)
        sortie(nT, 4) = nPresta
        sortie(nT, 5) = nPoint
        sortie(nT, 8) = IIf(txt(v(r, D_COMPTE)) = "0", 0, 1)
        sortie(nT, 9) = nIgn
        If enErreur Then
            sortie(nT, 7) = "Erreur"
            nErr = nErr + 1
        Else
            sortie(nT, 7) = "OK"
            For k = 1 To 6
                sortie(nT, 9 + k) = kmR(k)
            Next k
            If nPointe > 0 Then sortie(nT, 6) = Round(KmTotal(agent, dom, kmR), 2) Else sortie(nT, 6) = 0
        End If
        i = j + 1
    Loop
    EcrireTournees sortie, nT, joursCle, JoursPresents(v)

    If gNbManq > 0 Or nErr > 0 Or adrKo <> "" Then
        bilan = "Kilomètres : " & nT & " tournée(s) calculée(s) par la route (OpenStreetMap)."
        If nErr > 0 Then
            bilan = bilan & vbCrLf & nErr & " tournée(s) en erreur : " & _
                IIf(gLotsKo > 0, "service d'itinéraire OpenStreetMap injoignable", "itinéraire introuvable pour certains trajets") & _
                ". Relancer « Localiser les adresses et recalculer les km » (PARAMETRES)."
        End If
        If adrKo <> "" Then bilan = bilan & vbCrLf & "Adresses non localisées (erreur) : " & adrKo & "."
    End If
    CalculerKm = bilan
    Exit Function

Echec:
    CalculerKm = "Kilomètres non calculés (erreur) : " & Err.Description
    Application.StatusBar = False
End Function

' Total d'une tournée selon le départ et le retour (même règle que la carte)
'   avec maison : maison -> (agence le matin) -> chantiers -> (agence le soir) -> maison
'   sans maison : agence -> chantiers -> (agence le soir ou retour à l'agence)
Private Function KmTotal(ByVal agent As String, ByVal dom As String, km() As Variant) As Double
    Dim t As Double, matin As Boolean, soir As Boolean
    matin = Existe(gMatin, agent): soir = Existe(gSoir, agent)
    t = Nombre0(km(1))
    If dom <> "" Then
        If matin Then t = t + Nombre0(km(4)) + Nombre0(km(2)) Else t = t + Nombre0(km(3))
        If soir Then t = t + Nombre0(km(5)) + Nombre0(km(4)) Else t = t + Nombre0(km(6))
    Else
        t = t + Nombre0(km(2))
        If soir Or gRetour Then t = t + Nombre0(km(5))
    End If
    KmTotal = t
End Function

Private Function DepartDe(ByVal agent As String, ByVal dom As String) As String
    If dom = "" Then
        DepartDe = "Agence"
    ElseIf Existe(gMatin, agent) Then
        DepartDe = "Domicile + agence"
    Else
        DepartDe = "Domicile"
    End If
End Function

Private Function JoursPresents(v As Variant) As String
    Dim i As Long, s As String, d As Long
    s = "|"
    For i = 1 To UBound(v, 1)
        d = SerialDate(v(i, D_DATE))
        If d > 0 Then If InStr(s, "|" & d & "|") = 0 Then s = s & d & "|"
    Next i
    JoursPresents = s
End Function

Private Function LireParametres() As Boolean
    Dim ws As Worksheet, x As Variant
    Set ws = ThisWorkbook.Worksheets(F_PARAM)
    x = Nombre(ws.Range("H11").Value): If IsEmpty(x) Then Exit Function
    gAgLat = x
    x = Nombre(ws.Range("I11").Value): If IsEmpty(x) Then Exit Function
    gAgLon = x
    If gAgLat = 0 Or gAgLon = 0 Then Exit Function
    gRayon = Nombre0(ws.Range(P_RAYON).Value)
    If gRayon <= 0 Then gRayon = 60
    gRetour = OuiNon(ws.Range(P_RETOUR).Value, False)
    LireParametres = True
End Function

'--- Réglages des agents (PARAMETRES, à partir de la ligne P_AG1)
Private Sub ChargerAgents()
    Dim ws As Worksheet, i As Long, la As Double, lo As Double, nom As String
    Set gDom = New Collection: Set gMatin = New Collection: Set gSoir = New Collection
    Set ws = ThisWorkbook.Worksheets(F_PARAM)
    For i = P_AG1 To DerniereLigne(ws)
        nom = Trim$(txt(ws.Cells(i, 2).Value))
        If nom <> "" Then
            la = Nombre0(ws.Cells(i, 8).Value): lo = Nombre0(ws.Cells(i, 9).Value)
            If la <> 0 And lo <> 0 And OuiNon(ws.Cells(i, 5).Value, False) Then AjouterTexte gDom, nom, Pt(la, lo)
            If OuiNon(ws.Cells(i, 6).Value, False) Then AjouterTexte gMatin, nom, "1"
            If OuiNon(ws.Cells(i, 7).Value, False) Then AjouterTexte gSoir, nom, "1"
        End If
    Next i
End Sub

Private Function DomicileDe(ByVal nom As String) As String
    Dim x As Variant
    On Error GoTo Absent
    x = gDom.Item("k" & nom)
    If VarType(x) = vbString Then DomicileDe = x
    Exit Function
Absent:
    DomicileDe = ""
End Function

Private Sub AjouterTexte(c As Collection, ByVal cle As String, ByVal texte As String)
    On Error Resume Next
    c.Add texte, "k" & cle
End Sub

' Dernier indice du groupe (même jour, même agent) qui commence en i
Private Function FinGroupe(cles() As String, ByVal i As Long, ByVal n As Long) As Long
    Dim p As String, j As Long
    p = Left$(cles(i), InStrRev(cles(i), "|"))
    j = i
    Do While j < n
        If Left$(cles(j + 1), Len(p)) <> p Then Exit Do
        j = j + 1
    Loop
    FinGroupe = j
End Function

' Séquence des points pointés d'une tournée, dans l'ordre des clics
Private Sub Tournee(v As Variant, ix() As Long, ByVal i1 As Long, ByVal i2 As Long, _
                    pointe() As String, nPointe As Long, nPresta As Long, nPoint As Long, nIgn As Long)
    Dim i As Long, r As Long, la As Double, lo As Double, ok As Boolean, p As String
    Dim t() As Double, h1 As Variant, h2 As Variant, t1 As Double, t2 As Double

    ReDim pointe(1 To 2 * (i2 - i1 + 1))
    ReDim t(1 To 2 * (i2 - i1 + 1))
    nPointe = 0: nPresta = 0: nPoint = 0: nIgn = 0
    For i = i1 To i2
        r = ix(i)
        nPresta = nPresta + 1
        h1 = Heure(v(r, D_DEBR))
        If Not IsEmpty(h1) Then
            nPoint = nPoint + 1
            la = Nombre0(v(r, D_LAT)): lo = Nombre0(v(r, D_LON))
            ok = False
            If la <> 0 And lo <> 0 Then ok = (DistKm(gAgLat, gAgLon, la, lo) <= gRayon)
            If Not ok Then
                nIgn = nIgn + 1
            Else
                p = Pt(la, lo)
                t1 = CDbl(CLng(CDbl(h1) * 1440) Mod 1440)
                If UCase$(Trim$(txt(v(r, D_TYPE)))) = CONTENEURS Then
                    AjouterEtape pointe, t, nPointe, p, t1 * 10 + 1
                    h2 = Heure(v(r, D_FINR))
                    If Not IsEmpty(h2) Then
                        t2 = CDbl(CLng(CDbl(h2) * 1440) Mod 1440)
                        If t2 < t1 Then t2 = t2 + 1440
                        AjouterEtape pointe, t, nPointe, p, t2 * 10 + 2
                    End If
                Else
                    AjouterEtape pointe, t, nPointe, p, t1 * 10
                End If
            End If
        End If
    Next i
    OrdonnerEgalites pointe, t, nPointe
End Sub

' Clics à la même minute : on enchaîne au plus proche (même règle que la carte)
Private Sub OrdonnerEgalites(s() As String, t() As Double, ByVal n As Long)
    Dim i As Long, k As Long, a As Long, b As Long, best As Long, ps As String, pt2 As Double
    i = 1
    Do While i <= n
        k = i
        Do While k < n
            If Int(t(k + 1) / 10) <> Int(t(i) / 10) Then Exit Do
            k = k + 1
        Loop
        If k > i Then
            For a = i To k - 1
                If a > 1 Then
                    best = a
                    For b = a + 1 To k
                        If DistPts(s(a - 1), s(b)) < DistPts(s(a - 1), s(best)) Then best = b
                    Next b
                    If best > a Then
                        ps = s(best): pt2 = t(best)
                        For b = best To a + 1 Step -1
                            s(b) = s(b - 1): t(b) = t(b - 1)
                        Next b
                        s(a) = ps: t(a) = pt2
                    End If
                End If
            Next a
        End If
        i = k + 1
    Loop
End Sub

' Insertion triée (stable) d'une étape selon son heure de pointage
Private Sub AjouterEtape(s() As String, t() As Double, n As Long, ByVal p As String, ByVal cle As Double)
    Dim k As Long
    n = n + 1
    k = n
    Do While k > 1
        If t(k - 1) <= cle Then Exit Do
        s(k) = s(k - 1): t(k) = t(k - 1)
        k = k - 1
    Loop
    s(k) = p: t(k) = cle
End Sub

' Km d'une séquence : 1 trajet | 2 agence -> 1er | 3 domicile -> 1er | 4 domicile -> agence
'                     5 dernier -> agence | 6 dernier -> domicile   (Empty si sans objet)
Private Sub KmTournee(s() As String, ByVal n As Long, ByVal dom As String, km() As Variant, enErreur As Boolean)
    Dim k As Long, t As Double
    For k = 1 To 6
        km(k) = Empty
    Next k
    If n = 0 Then Exit Sub
    For k = 1 To n - 1
        t = t + KmT(s(k), s(k + 1), enErreur)
    Next k
    km(1) = Round(t, 2)
    km(2) = Round(KmT(gAg, s(1), enErreur), 2)
    km(5) = Round(KmT(s(n), gAg, enErreur), 2)
    If dom <> "" Then
        km(3) = Round(KmT(dom, s(1), enErreur), 2)
        km(4) = Round(KmT(dom, gAg, enErreur), 2)
        km(6) = Round(KmT(s(n), dom, enErreur), 2)
    End If
End Sub

' Distance routière connue ; sinon la tournée passe en erreur
Private Function KmT(ByVal a As String, ByVal b As String, enErreur As Boolean) As Double
    Dim km As Variant
    If a = b Then Exit Function
    km = LireCache(a & ">" & b)
    If IsEmpty(km) Then enErreur = True Else KmT = km
End Function

Private Sub BesoinsSequence(s() As String, ByVal n As Long, ByVal dom As String)
    Dim k As Long
    If n = 0 Then Exit Sub
    Besoin gAg, s(1)
    For k = 1 To n - 1
        Besoin s(k), s(k + 1)
    Next k
    Besoin s(n), gAg
    If dom <> "" Then
        Besoin dom, s(1)
        Besoin dom, gAg
        Besoin s(n), dom
    End If
End Sub

Private Sub Besoin(ByVal a As String, ByVal b As String)
    Dim cle As String
    If a = b Then Exit Sub
    cle = a & ">" & b
    If Not IsEmpty(LireCache(cle)) Then Exit Sub
    If Existe(gManq, cle) Then Exit Sub
    gNbManq = gNbManq + 1
    If gNbManq = 1 Then
        ReDim gManqA(1 To 256): ReDim gManqB(1 To 256)
    ElseIf gNbManq > UBound(gManqA) Then
        ReDim Preserve gManqA(1 To 2 * UBound(gManqA)): ReDim Preserve gManqB(1 To 2 * UBound(gManqB))
    End If
    gManqA(gNbManq) = a: gManqB(gNbManq) = b
    gManq.Add gNbManq, "k" & cle
End Sub

' Regroupe les trajets manquants en lots de MAX_PTS points, une requête par lot
Private Sub CalculerManquants()
    Dim k As Long, lot As Collection, pts() As String, np As Long
    Dim la() As Long, lb() As Long, nL As Long, ajoutes As Long
    ReDim pts(1 To MAX_PTS): ReDim la(1 To gNbManq): ReDim lb(1 To gNbManq)
    Set lot = New Collection
    For k = 1 To gNbManq
        If gHorsLigne Then Exit For
        ajoutes = 0
        If Idx(lot, gManqA(k)) = 0 Then ajoutes = ajoutes + 1
        If Idx(lot, gManqB(k)) = 0 And gManqB(k) <> gManqA(k) Then ajoutes = ajoutes + 1
        If np + ajoutes > MAX_PTS Then
            EnvoyerLot pts, np, la, lb, nL, k
            Set lot = New Collection: np = 0: nL = 0
        End If
        nL = nL + 1
        la(nL) = AjouterPoint(lot, pts, np, gManqA(k))
        lb(nL) = AjouterPoint(lot, pts, np, gManqB(k))
    Next k
    If nL > 0 And Not gHorsLigne Then EnvoyerLot pts, np, la, lb, nL, gNbManq + 1
    Application.StatusBar = False
End Sub

Private Function AjouterPoint(lot As Collection, pts() As String, np As Long, ByVal p As String) As Long
    AjouterPoint = Idx(lot, p)
    If AjouterPoint = 0 Then
        np = np + 1
        pts(np) = p
        lot.Add np, "k" & p
        AjouterPoint = np
    End If
End Function

' Envoie un lot (matrice de distances) ; kFin = indice du 1er trajet du lot suivant
Private Sub EnvoyerLot(pts() As String, ByVal np As Long, la() As Long, lb() As Long, ByVal nL As Long, ByVal kFin As Long)
    Dim url As String, rep As String, i As Long, c() As String, m() As Double, d As Double, k0 As Long
    gLots = gLots + 1
    Application.StatusBar = "Calcul des itinéraires : " & Format((kFin - 1) / gNbManq, "0%") & " (requête " & gLots & ")"
    DoEvents
    If gLots > 1 Then Pause 1.1                      ' service public : 1 requête par seconde
    For i = 1 To np
        c = Split(pts(i), ",")
        url = url & IIf(i > 1, ";", "") & c(1) & "," & c(0)          ' longitude,latitude
    Next i
    url = URL_ROUTE & url & "?annotations=distance"
    rep = HttpGet(url)
    If InStr(1, rep, """code"":""Ok""") = 0 Or Not LireMatrice(rep, np, m) Then
        gEchecs = gEchecs + 1: gLotsKo = gLotsKo + 1
        If gEchecs >= 2 Then gHorsLigne = True
        Exit Sub
    End If
    gEchecs = 0
    k0 = kFin - nL
    For i = 1 To nL
        d = m(la(i), lb(i))
        If d >= 0 Then AjouterCache gManqA(k0 + i - 1) & ">" & gManqB(k0 + i - 1), Round(d / 1000, 3)
    Next i
End Sub

' Lit la matrice "distances":[[..],[..]] d'une réponse OSRM (mètres ; null -> -1)
Private Function LireMatrice(ByVal json As String, ByVal n As Long, m() As Double) As Boolean
    Dim p As Long, a As Long, b As Long, i As Long, j As Long, t() As String, s As String
    ReDim m(1 To n, 1 To n)
    p = InStr(1, json, """distances"":[")
    If p = 0 Then Exit Function
    p = p + 13
    For i = 1 To n
        a = InStr(p, json, "[")
        If a = 0 Then Exit Function
        b = InStr(a + 1, json, "]")
        If b = 0 Then Exit Function
        t = Split(Mid$(json, a + 1, b - a - 1), ",")
        If UBound(t) + 1 <> n Then Exit Function
        For j = 1 To n
            s = Trim$(t(j - 1))
            If s = "null" Or s = "" Then m(i, j) = -1 Else m(i, j) = Val(s)
        Next j
        p = b + 1
    Next i
    LireMatrice = True
End Function

Private Function HttpGet(ByVal url As String) As String
    On Error Resume Next
    #If Mac Then
        HttpGet = MacScript("do shell script ""curl -s --max-time 30 '" & url & "'""")
    #Else
        Dim x As Object
        Set x = CreateObject("MSXML2.XMLHTTP.6.0")
        If x Is Nothing Then Set x = CreateObject("MSXML2.XMLHTTP")
        If x Is Nothing Then Exit Function
        x.Open "GET", url, False
        x.send
        If x.Status = 200 Then HttpGet = x.responseText
    #End If
End Function

Private Sub Pause(ByVal secondes As Double)
    Dim t0 As Double
    t0 = Timer
    Do While Timer >= t0 And Timer < t0 + secondes
        DoEvents
    Loop
End Sub

'--- Cache des distances (onglet masqué DISTANCES : trajet | km | calculé le)
Private Sub ChargerCache()
    Dim ws As Worksheet, lastR As Long, v As Variant, i As Long
    Set gCache = New Collection
    Set ws = ThisWorkbook.Worksheets(F_DIST)
    lastR = DerniereLigne(ws)
    If lastR < 2 Then Exit Sub
    v = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value
    On Error Resume Next
    For i = 1 To UBound(v, 1)
        If txt(v(i, 1)) <> "" And Not IsEmpty(v(i, 2)) And IsNumeric(v(i, 2)) Then gCache.Add CDbl(v(i, 2)), "k" & txt(v(i, 1))
    Next i
    On Error GoTo 0
End Sub

Private Function LireCache(ByVal cle As String) As Variant
    Dim x As Variant
    On Error GoTo Absent
    x = gCache.Item("k" & cle)
    If IsNumeric(x) Then LireCache = CDbl(x)
    Exit Function
Absent:
    LireCache = Empty
End Function

Private Sub AjouterCache(ByVal cle As String, ByVal km As Double)
    If Not IsEmpty(LireCache(cle)) Then Exit Sub
    gCache.Add km, "k" & cle
    gNbNouv = gNbNouv + 1
    If gNbNouv = 1 Then
        ReDim gNouvK(1 To 256): ReDim gNouvV(1 To 256)
    ElseIf gNbNouv > UBound(gNouvK) Then
        ReDim Preserve gNouvK(1 To 2 * UBound(gNouvK)): ReDim Preserve gNouvV(1 To 2 * UBound(gNouvV))
    End If
    gNouvK(gNbNouv) = cle: gNouvV(gNbNouv) = km
End Sub

Private Sub EcrireCache()
    Dim ws As Worksheet, arr As Variant, i As Long, depart As Long, maintenant As Date
    If gNbNouv = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(F_DIST)
    ReDim arr(1 To gNbNouv, 1 To 3)
    maintenant = Now
    For i = 1 To gNbNouv
        arr(i, 1) = gNouvK(i): arr(i, 2) = gNouvV(i): arr(i, 3) = maintenant
    Next i
    depart = DerniereLigne(ws) + 1
    ws.Cells(depart, 1).Resize(gNbNouv, 1).NumberFormat = "@"
    ws.Cells(depart, 1).Resize(gNbNouv, 3).Value = arr
    gNbNouv = 0
End Sub

'--- TOURNEES : on remplace les jours recalculés, on garde les autres jours encore présents
Private Sub EcrireTournees(nouv As Variant, ByVal nNouv As Long, ByVal joursCle As String, ByVal presents As String)
    Dim ws As Worksheet, lastR As Long, v As Variant, i As Long, j As Long, n As Long
    Dim tout As Variant, jour As Long, nAnc As Long
    Set ws = ThisWorkbook.Worksheets(F_TOURNEES)
    lastR = DerniereLigne(ws)
    If lastR >= T_LIGNE1 Then
        v = ws.Range(ws.Cells(T_LIGNE1, 1), ws.Cells(lastR, T_NBCOLS)).Value
        nAnc = UBound(v, 1)
    End If
    ReDim tout(1 To nAnc + nNouv + 1, 1 To T_NBCOLS)
    If joursCle <> "*" Then
        For i = 1 To nAnc
            If EstUneDate(v(i, 1)) And txt(v(i, 2)) <> "" Then
                jour = SerialDate(v(i, 1))
                If InStr(1, joursCle, "|" & jour & "|") = 0 And InStr(1, presents, "|" & jour & "|") > 0 Then
                    n = n + 1
                    For j = 1 To T_NBCOLS
                        tout(n, j) = v(i, j)
                    Next j
                    tout(n, 1) = CDate(jour)
                End If
            End If
        Next i
    End If
    For i = 1 To nNouv
        n = n + 1
        For j = 1 To T_NBCOLS
            tout(n, j) = nouv(i, j)
        Next j
    Next i
    If lastR >= T_LIGNE1 Then ws.Range(ws.Cells(T_LIGNE1, 1), ws.Cells(lastR, T_NBCOLS)).ClearContents
    ws.Range("A4").Resize(1, T_NBCOLS).Value = Array("Date", "Agent", "Départ", "Prestations", "Pointées", "Km pointés", _
        "Statut", "Compté", "Coord. ignorées", "Trajet entre chantiers", "Agence -> 1er", "Domicile -> 1er", _
        "Domicile -> agence", "Dernier -> agence", "Dernier -> domicile")
    If n = 0 Then Exit Sub
    With ws.Range(ws.Cells(T_LIGNE1, 1), ws.Cells(T_LIGNE1 + n - 1, T_NBCOLS))
        .Value = tout
        .Sort Key1:=.Cells(1, 1), Order1:=xlDescending, Key2:=.Cells(1, 2), Order2:=xlAscending, Header:=xlNo, MatchCase:=False
    End With
    ws.Range("A" & T_LIGNE1 & ":A" & (T_LIGNE1 + n - 1)).NumberFormat = "dd/mm/yyyy"
End Sub

'------------------------------------------------------------------------------
' Adresses (agence + domiciles) -> position GPS, via le géocodeur national
' (IGN Géoplateforme). Seule une adresse nouvelle ou modifiée est relocalisée.
'------------------------------------------------------------------------------
Private Function GeocoderAdresses() As String
    Dim ws As Worksheet, i As Long, lastR As Long, ko As String
    Set ws = ThisWorkbook.Worksheets(F_PARAM)
    If Not GeocoderLigne(ws, 11, 3) Then ko = "agence"
    lastR = DerniereLigne(ws)
    For i = P_AG1 To lastR
        If Trim$(txt(ws.Cells(i, 2).Value)) <> "" Then
            If Not GeocoderLigne(ws, i, 4) Then ko = ko & IIf(ko = "", "", ", ") & Trim$(txt(ws.Cells(i, 2).Value))
        End If
    Next i
    Application.StatusBar = False
    GeocoderAdresses = ko
End Function

' Adresse en colonne colAdr ; H latitude | I longitude | J (masquée) adresse localisée
Private Function GeocoderLigne(ws As Worksheet, ByVal ligne As Long, ByVal colAdr As Long) As Boolean
    Dim adr As String, la As Double, lo As Double, res As Long
    adr = Trim$(txt(ws.Cells(ligne, colAdr).Value))
    GeocoderLigne = True
    If adr = "" Then
        If Trim$(txt(ws.Cells(ligne, 10).Value)) <> "" Then ws.Range(ws.Cells(ligne, 8), ws.Cells(ligne, 10)).ClearContents
        Exit Function
    End If
    If adr = Trim$(txt(ws.Cells(ligne, 10).Value)) And Nombre0(ws.Cells(ligne, 8).Value) <> 0 Then Exit Function
    Application.StatusBar = "Localisation : " & adr
    res = Geocoder(adr, la, lo)
    If res = 1 Then
        ws.Cells(ligne, 8).Value = la
        ws.Cells(ligne, 9).Value = lo
        ws.Cells(ligne, 10).Value = adr
    Else
        ws.Cells(ligne, 8).Value = IIf(res = 0, "Introuvable", "Service injoignable")
        ws.Range(ws.Cells(ligne, 9), ws.Cells(ligne, 10)).ClearContents
        GeocoderLigne = False
    End If
End Function

' 1 = trouvée, 0 = introuvable (ou hors secteur), -1 = service injoignable
Private Function Geocoder(ByVal adr As String, la As Double, lo As Double) As Long
    Dim rep As String, p As Long, q As Long, c() As String
    rep = HttpGet("https://data.geopf.fr/geocodage/search?limit=1&lat=45.188&lon=5.724&q=" & UrlEncode(adr))
    If InStr(1, rep, """features""") = 0 Then Geocoder = -1: Exit Function
    p = InStr(1, rep, """coordinates"":[")
    If p = 0 Then Geocoder = 0: Exit Function
    p = p + 15
    q = InStr(p, rep, "]")
    If q = 0 Then Geocoder = 0: Exit Function
    c = Split(Mid$(rep, p, q - p), ",")
    If UBound(c) < 1 Then Geocoder = 0: Exit Function
    lo = Val(Trim$(c(0))): la = Val(Trim$(c(1)))
    If la = 0 Or lo = 0 Then Geocoder = 0: Exit Function
    If DistKm(45.188, 5.724, la, lo) > 80 Then Geocoder = 0: Exit Function     ' hors secteur grenoblois
    Geocoder = 1
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As Long, ch As String, o As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch): If c < 0 Then c = c + 65536
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Then
            o = o & ch
        ElseIf c < 128 Then
            o = o & "%" & Right$("0" & Hex$(c), 2)
        ElseIf c < 2048 Then
            o = o & "%" & Hex$(192 + (c \ 64)) & "%" & Hex$(128 + (c Mod 64))
        Else
            o = o & "%" & Hex$(224 + (c \ 4096)) & "%" & Hex$(128 + ((c \ 64) Mod 64)) & "%" & Hex$(128 + (c Mod 64))
        End If
    Next i
    UrlEncode = o
End Function

'--- Petits outils géographiques et JSON
' "45.17880,5.69230"
Private Function Pt(ByVal la As Double, ByVal lo As Double) As String
    Pt = Num(la) & "," & Num(lo)
End Function

Private Function DistPts(ByVal a As String, ByVal b As String) As Double
    Dim p() As String, q() As String
    p = Split(a, ","): q = Split(b, ",")
    DistPts = DistKm(Val(p(0)), Val(p(1)), Val(q(0)), Val(q(1)))
End Function

Private Function Iso(ByVal serial As Long) As String
    Dim d As Date
    d = CDate(serial)
    Iso = Year(d) & "-" & Format(Month(d), "00") & "-" & Format(Day(d), "00")
End Function

' Nombre au format JSON (point décimal, 5 décimales)
Private Function Num(ByVal x As Variant) As String
    Dim n As Variant
    n = Nombre(x)
    If IsEmpty(n) Then Num = "0": Exit Function
    Num = Replace(CStr(Round(CDbl(n), 5)), ",", ".")
End Function

' Chaîne JSON en ASCII pur (accents -> \uXXXX), sûre dans une balise <script>
Private Function JS(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    If s Like "*[!" & Chr$(32) & "-" & Chr$(126) & "]*" Or InStr(s, "<") > 0 Then
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1)
            code = AscW(ch)
            If code < 0 Then code = code + 65536
            If code < 32 Or code > 126 Or ch = "<" Then
                out = out & "\u" & Right$("000" & LCase$(Hex$(code)), 4)
            Else
                out = out & ch
            End If
        Next i
        s = out
    End If
    JS = """" & s & """"
End Function

'------------------------------------------------------------------------------
' Modifications faites sur la carte (agence matin / soir, adresse de la maison)
' La carte télécharge carte_cdamien.txt ; Excel le lit dans le dossier
' Téléchargements toutes les 3 secondes tant que la carte est ouverte (4 h max)
'------------------------------------------------------------------------------
Public Sub VeillerCarte()
    gVeille = False
    ImporterCarte
    If Now < gFinVeille Then ProgrammerVeille
End Sub

Private Sub DemarrerVeille()
    gFinVeille = Now + TimeSerial(4, 0, 0)
    If Not gVeille Then ProgrammerVeille
End Sub

Private Sub ProgrammerVeille()
    On Error Resume Next
    gProchain = Now + TimeSerial(0, 0, 3)
    Application.OnTime gProchain, "'" & ThisWorkbook.Name & "'!VeillerCarte"
    gVeille = (Err.Number = 0)
End Sub

Public Sub ArreterVeille()
    On Error Resume Next
    If gVeille Then Application.OnTime gProchain, "'" & ThisWorkbook.Name & "'!VeillerCarte", , False
    gVeille = False
End Sub

Public Sub ImporterCarte()
    Dim dossier As String, f As String, fichiers As New Collection, item As Variant, fic As Integer
    Dim ligne As String, c() As String, nb As Long, qui As String
    On Error GoTo Fin
    dossier = DossierTelechargements()
    If dossier = "" Then Exit Sub
    f = Dir(dossier & "carte_cdamien*.txt")
    Do While f <> ""
        fichiers.Add f
        f = Dir()
    Loop
    If fichiers.Count = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For Each item In fichiers
        fic = FreeFile
        Open dossier & CStr(item) For Input As #fic
        Do While Not EOF(fic)
            Line Input #fic, ligne
            c = Split(ligne, "|")
            If UBound(c) >= 7 Then
                If DecoderPct(c(0)) = "A" Then
                    AppliquerAgent DecoderPct(c(1)), Val(c(2)), Val(c(3)), (Val(c(4)) = 1), DecoderPct(c(5)), _
                                   Val(DecoderPct(c(6))), Val(DecoderPct(c(7)))
                    nb = nb + 1
                    qui = qui & IIf(qui = "", "", ", ") & DecoderPct(c(1))
                End If
            End If
        Loop
        Close #fic
        Kill dossier & CStr(item)
    Next item
    If nb > 0 Then
        MarquerAActualiser
        Application.StatusBar = "Carte enregistrée : " & qui & " (km recalculés à la prochaine mise à jour)"
    End If
Fin:
    On Error Resume Next
    Close #fic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Applique les choix de la carte pour un agent (PARAMETRES)
Private Sub AppliquerAgent(ByVal nom As String, ByVal matin As Long, ByVal soir As Long, ByVal adrModifiee As Boolean, _
                           ByVal adr As String, ByVal la As Double, ByVal lo As Double)
    Dim wsP As Worksheet, r As Long
    nom = Trim$(nom)
    If nom = "" Then Exit Sub
    Set wsP = ThisWorkbook.Worksheets(F_PARAM)
    r = P_AG1
    Do While Trim$(txt(wsP.Cells(r, 2).Value)) <> ""
        If Trim$(txt(wsP.Cells(r, 2).Value)) = nom Then Exit Do
        r = r + 1
    Loop
    wsP.Cells(r, 2).Value = nom
    If Trim$(txt(wsP.Cells(r, 3).Value)) = "" Then wsP.Cells(r, 3).Value = "Oui"
    wsP.Cells(r, 6).Value = IIf(matin = 1, "Oui", "Non")
    wsP.Cells(r, 7).Value = IIf(soir = 1, "Oui", "Non")
    If Not adrModifiee Then Exit Sub
    If Trim$(adr) = "" Then
        wsP.Cells(r, 4).ClearContents
        wsP.Cells(r, 5).Value = "Non"
        wsP.Range(wsP.Cells(r, 8), wsP.Cells(r, 10)).ClearContents
    Else
        wsP.Cells(r, 4).Value = Trim$(adr)
        wsP.Cells(r, 5).Value = "Oui"
        If la <> 0 And lo <> 0 Then
            wsP.Cells(r, 8).Value = la
            wsP.Cells(r, 9).Value = lo
            wsP.Cells(r, 10).Value = Trim$(adr)
        Else
            wsP.Range(wsP.Cells(r, 8), wsP.Cells(r, 10)).ClearContents     ' localisée au calcul des km
        End If
    End If
End Sub

Private Function DossierTelechargements() As String
    Dim d As String
    On Error Resume Next
    #If Mac Then
        d = Environ("HOME") & "/Downloads/"
    #Else
        d = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.Path
        If d = "" Then d = Environ("USERPROFILE") & "\Downloads"
        If Right$(d, 1) <> "\" Then d = d & "\"
    #End If
    DossierTelechargements = d
End Function

' Texte encodé en %XX (UTF-8) -> texte
Private Function DecoderPct(ByVal s As String) As String
    Dim i As Long, b() As Long, n As Long, o As String, c As Long, k As Long, cp As Long
    If s = "" Then Exit Function
    ReDim b(1 To Len(s) + 3)
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) = "%" And i + 2 <= Len(s) Then
            n = n + 1: b(n) = CLng("&H" & Mid$(s, i + 1, 2)): i = i + 3
        Else
            n = n + 1: b(n) = AscW(Mid$(s, i, 1)) And &HFF: i = i + 1
        End If
    Loop
    k = 1
    Do While k <= n
        c = b(k)
        If c < &H80 Then
            cp = c: k = k + 1
        ElseIf c < &HE0 And k + 1 <= n Then
            cp = ((c And &H1F) * 64) Or (b(k + 1) And &H3F): k = k + 2
        ElseIf c < &HF0 And k + 2 <= n Then
            cp = ((c And &HF) * 4096) Or ((b(k + 1) And &H3F) * 64) Or (b(k + 2) And &H3F): k = k + 3
        Else
            cp = &HFFFD: k = k + 4
        End If
        o = o & ChrW(cp)
    Loop
    DecoderPct = o
End Function
