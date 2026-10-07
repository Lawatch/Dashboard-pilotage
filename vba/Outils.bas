'==============================================================================
' Outils communs : noms d'onglets, lecture de cellules, collections, tris, mise en forme
'==============================================================================
Option Explicit

Public Const F_DASH As String = "DASHBOARD"
Public Const F_CTRL As String = "CONTROLES"
Public Const F_PARAM As String = "PARAMETRES"
Public Const F_ABO As String = "ABONNEMENTS"
Public Const F_DATA As String = "DATA"
Public Const F_ABOS As String = "ABOS"
Public Const F_LISTES As String = "LISTES"
Public Const F_TOURNEES As String = "TOURNEES"
Public Const F_DIST As String = "DISTANCES"
Public Const F_MODELE As String = "MODELE_CARTE"

' PARAMETRES
Public Const P_AG1 As Long = 20                 ' 1re ligne du tableau des agents (B:J)
Public Const P_MAN1 As Long = 20                ' 1re ligne des rattachements manuels (L:N)
Public Const P_ADR_AGENCE As String = "C11"
Public Const P_RETOUR As String = "C12"
Public Const P_RAYON As String = "C13"
Public Const P_JOURS_REF As String = "C14"

' DATA : une ligne par prestation (colonnes 1 à 17 lues aussi par la carte)
Public Const NB_COLS As Long = 24
Public Const D_DATE As Long = 1, D_REF As Long = 2, D_AGENT As Long = 3, D_PROJET As Long = 4
Public Const D_CHANTIER As Long = 5, D_LIBELLE As Long = 6, D_CAT As Long = 7, D_TYPE As Long = 8
Public Const D_DEBP As Long = 9, D_FINP As Long = 10, D_DUREEP As Long = 11, D_DEBR As Long = 12
Public Const D_FINR As Long = 13, D_DUREER As Long = 14, D_POINTEE As Long = 15, D_LAT As Long = 16
Public Const D_LON As Long = 17, D_ONGLET As Long = 18, D_COMPTE As Long = 19, D_ABO As Long = 20
Public Const D_MODE As Long = 21, D_CLE As Long = 22, D_HRET As Long = 23, D_STATUT As Long = 24

Public Const CONTENEURS As String = "GESTCONT"   ' 1er clic = sortie, 2e clic = rentrée

'------------------------------------------------------------------------------
' Onglets
'------------------------------------------------------------------------------
Public Function EstOngletSysteme(ByVal nom As String) As Boolean
    Select Case UCase$(nom)
        Case F_DASH, F_CTRL, F_PARAM, F_ABO, F_DATA, F_ABOS, F_LISTES, F_TOURNEES, F_DIST, F_MODELE
            EstOngletSysteme = True
    End Select
End Function

' Onglet mis de côté (doublon / ancienne version d'un jour)
Public Function EstOngletEcarte(ByVal nom As String) As Boolean
    Dim n As String
    n = LCase$(nom)
    EstOngletEcarte = (InStr(n, "doublon") > 0 Or InStr(n, "ancien") > 0)
End Function

Public Function ExisteOnglet(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nom, vbTextCompare) = 0 Then ExisteOnglet = True: Exit Function
    Next ws
End Function

Public Function NomLibre(ByVal racine As String) As String
    Dim i As Long
    NomLibre = Left$(racine, 31)
    i = 2
    Do While ExisteOnglet(NomLibre)
        NomLibre = Left$(racine, 26) & " (" & i & ")"
        i = i + 1
    Loop
End Function

' 46301 -> "061026"
Public Function JJMMAA(ByVal serial As Long) As String
    Dim d As Date
    d = CDate(serial)
    JJMMAA = Format(Day(d), "00") & Format(Month(d), "00") & Format(Year(d) Mod 100, "00")
End Function

' "061026" -> numéro de série (0 si invalide)
Public Function SerialDepuisNom(ByVal s As String) As Long
    Dim j As Long, m As Long, a As Long, d As Date
    If Not s Like "######" Then Exit Function
    j = CLng(Left$(s, 2)): m = CLng(Mid$(s, 3, 2)): a = 2000 + CLng(Right$(s, 2))
    If m < 1 Or m > 12 Or j < 1 Or j > 31 Then Exit Function
    d = DateSerial(a, m, j)
    If Day(d) <> j Then Exit Function
    SerialDepuisNom = CLng(d)
End Function

Public Sub ColorerOnglet(ws As Worksheet, ByVal couleur As Long)
    On Error Resume Next
    If couleur < 0 Then ws.Tab.ColorIndex = xlColorIndexNone Else ws.Tab.Color = couleur
    On Error GoTo 0
End Sub

Public Function DerniereLigne(ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If r Is Nothing Then DerniereLigne = 1 Else DerniereLigne = r.Row
End Function

Public Function DerniereColonne(ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If r Is Nothing Then DerniereColonne = 1 Else DerniereColonne = r.Column
End Function

'------------------------------------------------------------------------------
' Lecture de cellules
'------------------------------------------------------------------------------
' Texte sûr (les cellules en erreur deviennent "")
Public Function txt(ByVal x As Variant) As String
    If IsError(x) Or IsEmpty(x) Or IsNull(x) Then Exit Function
    txt = CStr(x)
End Function

Public Function EstUneDate(ByVal x As Variant) As Boolean
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If VarType(x) = vbDate Then EstUneDate = True: Exit Function
    If VarType(x) = vbString Then
        If Len(x) >= 8 And x Like "*#/*#/##*" Then EstUneDate = IsDate(x)
        Exit Function
    End If
    If IsNumeric(x) Then EstUneDate = (CDbl(x) > 30000 And CDbl(x) < 80000)
End Function

Public Function SerialDate(ByVal x As Variant) As Long
    If EstUneDate(x) Then SerialDate = CLng(Int(CDbl(CDate(x))))
End Function

' Nombre depuis une cellule (gère "1 029,63", espaces insécables, "45.33") ; Empty si vide
Public Function Nombre(ByVal x As Variant) As Variant
    Dim s As String
    Nombre = Empty
    If IsError(x) Or IsEmpty(x) Then Exit Function
    Select Case VarType(x)
        Case vbDouble, vbLong, vbInteger, vbCurrency, vbSingle, vbDecimal, vbByte
            Nombre = CDbl(x): Exit Function
    End Select
    s = Replace(Replace(Replace(Trim$(CStr(x)), ChrW(160), ""), ChrW(8239), ""), " ", "")
    s = Replace(s, ",", ".")
    If s Like "*[0-9]*" Then Nombre = Val(s)
End Function

Public Function Nombre0(ByVal x As Variant) As Double
    Dim y As Variant
    y = Nombre(x)
    If Not IsEmpty(y) Then Nombre0 = y
End Function

' Heure depuis une cellule (valeur heure Excel ou texte "hh:mm") -> fraction de jour ; Empty si vide
Public Function Heure(ByVal x As Variant) As Variant
    Dim s As String, p() As String
    Heure = Empty
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If VarType(x) = vbDate Or VarType(x) = vbDouble Then
        Heure = CDbl(x) - Int(CDbl(x)): Exit Function
    End If
    s = Trim$(CStr(x))
    If s Like "#*:##*" Then
        p = Split(s, ":")
        Heure = (Val(p(0)) * 60 + Val(p(1))) / 1440
    End If
End Function

Public Function HHMM(ByVal x As Variant) As String
    Dim h As Variant, m As Long
    h = Heure(x)
    If IsEmpty(h) Then Exit Function
    m = CLng(h * 1440) Mod 1440
    HHMM = Format(m \ 60, "00") & ":" & Format(m Mod 60, "00")
End Function

' Minutes depuis minuit (0 si vide)
Public Function Minutes(ByVal x As Variant) As Long
    Dim h As Variant
    h = Heure(x)
    If Not IsEmpty(h) Then Minutes = CLng(CDbl(h) * 1440) Mod 1440
End Function

' Nom normalisé : majuscules sans accents, lettres et chiffres séparés par un espace
Public Function Norm(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, o As String, avecAcc As String, sansAcc As String, p As Long, blanc As Boolean
    avecAcc = ChrW(192) & ChrW(193) & ChrW(194) & ChrW(195) & ChrW(196) & ChrW(197) & ChrW(199) & ChrW(200) & ChrW(201) & ChrW(202) & ChrW(203) & _
              ChrW(204) & ChrW(205) & ChrW(206) & ChrW(207) & ChrW(209) & ChrW(210) & ChrW(211) & ChrW(212) & ChrW(213) & ChrW(214) & _
              ChrW(217) & ChrW(218) & ChrW(219) & ChrW(220) & ChrW(221) & ChrW(376) & ChrW(338) & ChrW(198)
    sansAcc = "AAAAAACEEEEIIIINOOOOOUUUUYYOA"
    s = UCase$(s)
    blanc = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        p = InStr(1, avecAcc, ch, vbBinaryCompare)
        If p > 0 Then ch = Mid$(sansAcc, p, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Then
            o = o & ch: blanc = False
        ElseIf Not blanc Then
            o = o & " ": blanc = True
        End If
    Next i
    Norm = Trim$(o)
End Function

' Agent fictif (poste à pourvoir, tournée de remplacement...) : ses prestations ne sont pas affectées
Public Function EstAgentFictif(ByVal nom As String) As Boolean
    Static cache As Collection
    Dim n As String, r As Boolean
    If cache Is Nothing Then Set cache = New Collection
    If Existe(cache, nom) Then EstAgentFictif = (cache.Item("k" & nom) = 1): Exit Function
    n = Norm(nom)
    r = (InStr(n, "SUPPRIMER") > 0 Or InStr(n, "REMPLACEMENT") > 0 Or InStr(n, "DEFINIR") > 0 _
         Or InStr(n, "NOUVEAU") > 0 Or n = "")
    cache.Add IIf(r, 1, 0), "k" & nom
    EstAgentFictif = r
End Function

' Norm avec mémoire (les mêmes noms de chantiers reviennent sur des milliers de lignes)
Public Function NormC(ByVal s As String) As String
    Static cache As Collection, n As Long
    If cache Is Nothing Or n > 20000 Then Set cache = New Collection: n = 0
    If Existe(cache, s) Then NormC = cache.Item("k" & s): Exit Function
    NormC = Norm(s)
    cache.Add NormC, "k" & s
    n = n + 1
End Function

' Distance à vol d'oiseau (km, formule de haversine)
Public Function DistKm(ByVal la1 As Double, ByVal lo1 As Double, ByVal la2 As Double, ByVal lo2 As Double) As Double
    Dim r As Double, h As Double
    r = 3.14159265358979 / 180
    h = Sin((la2 - la1) * r / 2) ^ 2 + Cos(la1 * r) * Cos(la2 * r) * Sin((lo2 - lo1) * r / 2) ^ 2
    If h <= 0 Then Exit Function
    If h >= 1 Then DistKm = 20000: Exit Function
    DistKm = 2 * 6371 * Atn(Sqr(h) / Sqr(1 - h))
End Function

Public Function OuiNon(ByVal x As Variant, ByVal defaut As Boolean) As Boolean
    Dim s As String
    s = LCase$(Trim$(txt(x)))
    If s = "" Then OuiNon = defaut: Exit Function
    OuiNon = (s = "oui" Or s = "1" Or s = "x" Or s = "vrai" Or s = ChrW(9745))
End Function

'------------------------------------------------------------------------------
' Collections (clé texte -> valeur numérique)
'------------------------------------------------------------------------------
Public Sub Ajout(c As Collection, ByVal cle As String, ByVal x As Variant)
    On Error Resume Next
    c.Add x, "k" & cle
    On Error GoTo 0
End Sub

' Valeur mémorisée (0 si absente)
Public Function Idx(c As Collection, ByVal cle As String) As Long
    On Error GoTo Absent
    Idx = CLng(c.Item("k" & cle))
    Exit Function
Absent:
    Idx = 0
End Function

Public Function Existe(c As Collection, ByVal cle As String) As Boolean
    Dim x As Variant
    On Error GoTo Absent
    x = c.Item("k" & cle)
    Existe = True
    Exit Function
Absent:
    Existe = False
End Function

Public Function Valeur(c As Collection, ByVal cle As String) As Variant
    On Error GoTo Absent
    Valeur = c.Item("k" & cle)
    Exit Function
Absent:
    Valeur = Empty
End Function

' Ajoute la clé si elle est nouvelle ; True si ajoutée
Public Function AjoutUnique(c As Collection, ByVal cle As String) As Boolean
    On Error GoTo Deja
    c.Add 1, "k" & cle
    AjoutUnique = True
    Exit Function
Deja:
    AjoutUnique = False
End Function

'------------------------------------------------------------------------------
' Tris (tri rapide avec tableau d'indices)
'------------------------------------------------------------------------------
Public Sub TriTexte(cles() As String, ix() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, ts As String, ti As Long
    If hi <= lo Then Exit Sub
    i = lo: j = hi
    p = cles((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(cles(i), p, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(cles(j), p, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            ts = cles(i): cles(i) = cles(j): cles(j) = ts
            ti = ix(i): ix(i) = ix(j): ix(j) = ti
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then TriTexte cles, ix, lo, j
    If i < hi Then TriTexte cles, ix, i, hi
End Sub

Public Sub TriDesc(vals() As Double, ix() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, tv As Double, ti As Long
    If hi <= lo Then Exit Sub
    i = lo: j = hi
    p = vals((lo + hi) \ 2)
    Do While i <= j
        Do While vals(i) > p
            i = i + 1
        Loop
        Do While vals(j) < p
            j = j - 1
        Loop
        If i <= j Then
            tv = vals(i): vals(i) = vals(j): vals(j) = tv
            ti = ix(i): ix(i) = ix(j): ix(j) = ti
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then TriDesc vals, ix, lo, j
    If i < hi Then TriDesc vals, ix, i, hi
End Sub

'------------------------------------------------------------------------------
' Mise en forme des tableaux écrits par les macros
'------------------------------------------------------------------------------
Public Sub FormatTitre(r As Range, ByVal couleur As Long)
    With r
        .Interior.Color = couleur
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .RowHeight = 20
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub FormatEntete(r As Range)
    With r
        .Interior.Color = RGB(231, 236, 242)
        .Font.Bold = True
        .Font.Color = RGB(31, 58, 95)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(31, 58, 95)
    End With
End Sub

Public Sub FormatAide(r As Range)
    With r
        .Font.Italic = True
        .Font.Color = RGB(110, 110, 110)
        .Font.Size = 9
    End With
End Sub

Public Sub FormatLignes(r As Range)
    With r.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(225, 225, 225)
    End With
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(225, 225, 225)
    End With
End Sub

' Cellule cliquable (lien vers elle-même, intercepté par Worksheet_FollowHyperlink)
Public Sub Cliquable(c As Range, ByVal texte As String)
    c.Value = texte
    Lien c, "'" & c.Worksheet.Name & "'!" & c.Address(False, False), texte
End Sub

' Lien interne (sans effet si l'application ne sait pas le créer)
Public Sub Lien(c As Range, ByVal cible As String, ByVal texte As String)
    On Error Resume Next
    c.Worksheet.Hyperlinks.Add Anchor:=c, Address:="", SubAddress:=cible, TextToDisplay:=texte
    If Err.Number <> 0 Then c.Value = texte
    On Error GoTo 0
End Sub

Public Sub EffacerZone(r As Range)
    On Error Resume Next
    r.Hyperlinks.Delete
    r.UnMerge
    r.Clear
    r.RowHeight = 12.75                          ' hauteur standard (Arial 10)
    On Error GoTo 0
End Sub

Public Function Pct(ByVal a As Double, ByVal b As Double) As Variant
    If b = 0 Then Pct = "" Else Pct = a / b
End Function

Public Function TxtPeriode(ByVal d1 As Double, ByVal d2 As Double) As String
    If d1 = d2 Then
        TxtPeriode = Format(d1, "dd/mm/yyyy")
    Else
        TxtPeriode = Format(d1, "dd/mm") & " au " & Format(d2, "dd/mm/yyyy")
    End If
End Function

Public Function Euros(ByVal x As Double) As String
    Euros = Format(x, "#,##0") & " " & ChrW(8364)
End Function
