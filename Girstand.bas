Attribute VB_Name = "GirStand"
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module standard : GirStand   -   GIRSTAND.BAS
'   Module comportant différentes petites fonctions utilitaires

'**************************************************************************************

Option Explicit

'**************************************************
' ArcCosinus
'**************************************************
Public Function Arccos(ByVal X As Double) As Double
  Select Case X
  Case 1
    Arccos = 0
  Case -1
    Arccos = pi
  Case Else
    Arccos = Atn(-X / Sqr(-X * X + 1)) + pi / 2
  End Select
  
End Function

'****************************************************************
' Conversion d'un angle d'unité usuelle en radian ou inversement
'******************************************************************
Public Function angConv(ByVal Angle As Double, ByVal Enradian As Boolean) As Double
'  If radian Then angConv = Angle * pi / 180 Else angConv = Angle * 180 / pi
' eqvPI vaut 180 po 200 selon que les unites sont en degrés  ou en grade

  If Enradian Then angConv = Angle * pi / eqvPI(OptGen.unite) Else angConv = Angle * eqvPI(OptGen.unite) / pi

End Function

Public Sub ExecuteZoomEtendu()

  verifRetaillage
  If FichierJournal Then Write #numFichLog, "Tout Voir"
  ToutVoir
  
  TerminerRecadrage

End Sub

'**************************************************************************************
' Translation-Rotation du point d'insertion d'un bloc avec ou sans facteur d'échelle
'**************************************************************************************
Public Function TransRot(ByVal p As PT, ByVal Trans As PT, ByVal Alpha As Single, ByVal Echelle As Single, Optional ByVal EchelleY As Single) As PT
Dim p0 As New PT

  ' Facteur d'échelle
  If EchelleY = 0 Then EchelleY = Echelle
  p0.X = p.X * Echelle
  p0.Y = p.Y * EchelleY
  ' Rotation
  If Alpha <> 0 Then
    Set p0 = Rotation(p0, angConv(Alpha, radian))
  End If
  ' translation
  p0.X = p0.X + Trans.X
  p0.Y = p0.Y + Trans.Y
  
  Set TransRot = p0
  
End Function

'*********1*****************************************************************************
' Rotation
'**************************************************************************************
Public Function Rotation(ByVal p As PT, ByVal Alpha As Single) As PT
Dim p0 As New PT

  p0.X = p.X * Cos(Alpha) - p.Y * Sin(Alpha)
  p0.Y = p.X * Sin(Alpha) + p.Y * Cos(Alpha)
  Set Rotation = p0
  
End Function

'*********1*****************************************************************************
' Rotation d'un point suivie d'une translation
'**************************************************************************************
Public Function RotTrans(ByVal p As PT, ByVal Trans As PT, ByVal Alpha As Single) As PT
Dim p0 As New PT
  
  Set p0 = Rotation(p, Alpha)
  p0.X = p0.X + Trans.X
  p0.Y = p0.Y + Trans.Y
  Set RotTrans = p0
  
End Function

'*********1*****************************************************************************
' Milieu de 2 points
'**************************************************************************************
Public Function pMilieu(ByVal p1 As PT, ByVal p2 As PT) As PT
Dim p0 As New PT
    
    p0.X = (p1.X + p2.X) / 2
    p0.Y = (p1.Y + p2.Y) / 2
    Set pMilieu = p0
    
End Function

'*********1*****************************************************************************
' Calcul de l'angle formé par 2 points ordonnés p1 et p2
' Retourne un angle compris entre ]-pi,pi]
'**************************************************************************************
Public Function CalcAngle(ByVal p1 As PT, ByVal p2 As PT) As Double

  If p1.X <> p2.X Then
    CalcAngle = Atn((p2.Y - p1.Y) / (p2.X - p1.X))
    If p2.X < p1.X Then
    ' L'angle déterminé ci-dessus appartient à ]-pi/2,pi/2[
    'Il appartient en fait à ]-pi,-pi/2[ ou à ]pi/2,pi]
      Select Case Sgn(CalcAngle)
      Case 1
        CalcAngle = CalcAngle - pi
      Case Else ' 0 ou -1
        CalcAngle = CalcAngle + pi
      End Select
    End If
    
  Else
      ' Droite verticale
    If p2.Y > p1.Y Then
      CalcAngle = pi / 2
    Else
      CalcAngle = -pi / 2
    End If
  End If

End Function

'*********1*****************************************************************************
' Distance entre 2 points
'**************************************************************************************
Public Function Distance(ByVal X As Double, ByVal X1 As Double, ByVal Y As Double, ByVal Y1 As Double) As Double

  Distance = Sqr(Carre(X1 - X) + Carre(Y1 - Y))
  
End Function

'*********1*****************************************************************************
'Ecriture d'un texte en bas à gauche d'un objet (Picture, Form...)
'**************************************************************************************
Public Sub bgTexte(ByVal Msg As String)
Dim sauv As Boolean
Static lgTexte As Integer
Dim pct As VB.PictureBox
  
  Set pct = ObjetDessin
  
  With pct
    .CurrentX = 0  '.TextWidth(Msg)
    .CurrentY = .Height - 1.5 * .TextHeight(Msg)
    sauv = .FontTransparent
    .Font.Italic = True
    .Font.Bold = True
    .FontTransparent = False
    If Msg = "" Then
      Msg = String(lgTexte * 1.5, " ")
    Else
      lgTexte = Len(Msg) ' .TextWidth(Msg) / .TextWidth("E")
    End If
  End With
  
  pct.Print Msg
  
  With pct
    .FontTransparent = sauv
    .Font.Italic = False
    .Font.Bold = False
  End With
  
End Sub

'*********1*****************************************************************************
'Arrondi d'un nombre avec n décimales + une éventuelle précision (à 'm' près)
' si precision=-1, indique qu'il faut l'ignorer
'**************************************************************************************
'Public Function Arrond(ByVal v As Single, ByVal Nbdec As Integer, Optional ByVal precision As Integer = -1) As String
' CERTU/ESi/GNSI le 10/01/2014 changement du type Single en Double par cohérence
' par exemple la valeur 999999,75 perdait ses décimales en Single

Public Function Arrond(ByVal v As Double, ByVal Nbdec As Integer, Optional ByVal precision As Integer = -1) As String
Dim fm As String      ' Format
Dim w As Double
Dim strPtDecimal As String

'Mise en Commentaire (AV : 6/5/03 : Suppression des Variant
'De + ce controle ne sert à rien car il est éventuellement fait en amont
'If Not IsNumeric(v) Then Arrond = 0: Exit Function

  If Abs(Nbdec) = 0 Then
    fm = "0"
  Else
    fm = "0." & String(Abs(Nbdec), "0")
    If precision <> -1 Then    ' arrondi à la précision supérieure  (ex. si precision=5 et Nbdec=3        2.4524 sera arrondi à 2.455
       w = 10 ^ Nbdec
      v = Fix((v * w + precision - 0.1) / precision) * precision / w
    End If
  End If
   
  Arrond = Format(v, fm)
  strPtDecimal = Chr(gtPtDecimal)
  
' Elimination des zéros non significatifs après le point décimal
  If InStr(Arrond, strPtDecimal) Then
    Do While Right(Arrond, 1) = "0"
      Arrond = Left(Arrond, Len(Arrond) - 1)
    Loop
    ' suppression du Pt décimal s'il n'y a plus rien derrière
    If Right(Arrond, 1) = strPtDecimal Then Arrond = Left(Arrond, Len(Arrond) - 1)
  End If
  
End Function

Public Function Min(ParamArray a()) As Double
Dim i As Integer

  Min = a(0)
  For i = 1 To UBound(a)
    If a(i) < Min Then Min = a(i)
  Next
End Function

Public Function Max(ParamArray a()) As Double
Dim i As Integer

  Max = a(0)
  For i = 1 To UBound(a)
    If a(i) > Max Then Max = a(i)
  Next
End Function

Public Function Superieur(ByVal v1 As Double, ByVal v2 As Double) As Boolean

  Superieur = (v1 > v2 + 0.000005)
  
End Function

Public Function Carre(ByVal v As Double) As Double
  Carre = v ^ 2
End Function

Private Function ordre_idee(ByVal nb As Double) As Integer
  
' retourne l'ordre d'idée d'un nombre
'ex: 783=7.83*10² --> ordre_idee=2

  If nb >= 10 Then
    ordre_idee = ordre_idee(nb / 10) + 1
  Else
    ordre_idee = 0
  End If
  
End Function

Public Function superFix(ByVal nb As Double, ByVal Nbchiffre As Integer) As String
' retourne au maximum les derniers Nbchiffre de Nb
' ex: Si Nb=65984 et Nbchiffre=3 --> superFix=984
'       Si Nb=4589   et Nbchiffre=4 --> superFix=4589

Dim i As Integer
Dim fm As String

  i = ordre_idee(Abs(nb))
  If i >= Nbchiffre Then
    superFix = nb - Fix(nb / 10 ^ Nbchiffre) * 10 ^ Nbchiffre
    fm = "\.\.\."
  Else
    superFix = nb
    fm = ""
  End If
  
  For i = 1 To Nbchiffre - 1
    fm = fm & "#"
  Next
  fm = fm & "0"
  
  superFix = Format(superFix, fm)

End Function

Public Function decompRGB(ByVal v As Long) As String
' Décomposition d'une couleur RGB en ses 3 composants sous la forme "<compos1>[ <compos2>[ <compos3>]]"
Dim q As Long, reste As Long

  If v < 256 Then
    decompRGB = CStr(v)
  Else
    q = Int(v / 256)
    reste = v - q * 256
    decompRGB = decompRGB(reste) & " " & decompRGB(q)
  End If
    
End Function

Public Function recompRGB(ByVal v As String) As Long
' Recomposition d'une couleur RGB à partir de ses 3 composants (séparés par des blancs)
  Dim i%, n%, deb%, fin%
  
  ' v = suppCNull(v)        ' cette fonction était appelée suite au retour d'une DLL C (GetProfileString) avant l'utilisation de la Registry
  deb% = 1
  
  fin = InStr(v, " ")
  While fin <> 0
    recompRGB = recompRGB + val(Mid(v, deb, fin - deb)) * 256 ^ n%
    n = n + 1
    deb = fin + 1
    fin = InStr(deb, v, " ")
  Wend
  recompRGB = recompRGB + val(Mid(v, deb, Len(v) - deb + 1)) * 256 ^ n%
    
End Function

Public Function valChaine(ByVal chaine As String) As Double
' Transforme en numérique une chaine C

On Error GoTo GestErr
  valChaine = CDbl(suppCNull(chaine))
  Exit Function
  
GestErr:
  ' L'utilisateur a pu changer de config au niveau du séparateur décimal
'  valChaine = valChaine(substVirgulePoint(chaine))
  valChaine = valChaine(substPtDecimalRegional(chaine, Regional:=True))
Exit Function
  
End Function

Public Function suppCNull(v As String) As String
' Supprime tous les caractères après et y compris le  caractère NULL d'une chaine C

  suppCNull = Left(v, InStr(v, Chr(0)) - 1)
  
End Function

Public Function Majus(ByVal Texte As String, Optional ByVal Tout) As String
  If IsMissing(Tout) Then
    Majus = UCase(Left(Texte, 1)) & Mid(Texte, 2)
  Else
    Majus = UCase(Texte)
  End If
End Function

Public Function Minus(ByVal Texte As String, Optional ByVal Tout) As String
  If IsMissing(Tout) Then
    Minus = LCase(Left(Texte, 1)) & Mid(Texte, 2)
  Else
    Minus = LCase(Texte)
  End If
End Function

Public Function Numopt(ByVal bouton As Object) As Integer
' Retourne le numéro d'option en cours dans un groupe de boutons
' Retourne -1 si aucun
  Dim i%

  Numopt = -1   ' a priori aucun bouton sélectionné
  For i = 1 To bouton.Count
    If bouton(i - 1) = True Then Numopt = i - 1: Exit Function  ' bouton trouvé
  Next
  
End Function

'*************************************************************************************
' Existence d'un fichier
'   Nom : nom du fichier
'   Retourne True si le fichier existe
'*************************************************************************************
Public Function ExistFich(ByVal Nom As String, Optional ByVal attrib As Integer) As Boolean
  
  If attrib = 0 Then
    ExistFich = gtFso.FileExists(Nom)
  Else
    ExistFich = gtFso.FolderExists(Nom)
  End If
  
End Function

'*************************************************************************
' Recherche de l'existence d'un véhicule
' L'appel implicite de la méthode Item génère une erreur si le véhicule n'existe pas
'*************************************************************************
Public Function ExistVeh(ByVal Nom As String, ByVal ColVeh As Collection) As Boolean
Dim objet As Object

On Error GoTo GestErr

  Set objet = ColVeh(UCase(Nom))
  ExistVeh = True
Exit Function
  
GestErr:
  If Err = 5 Then Exit Function Else ErrGeneral
    
End Function

Public Function extraiRep(ByVal s As String) As String
' Extrait le chemin d'un nom de fichier, y compris l'éventuel dernier '\'
  Dim pos%
  
  extraiRep = gtFso.GetParentFolderName(s)
  Exit Function
  
    pos = InStr(s, "\")
    If pos <> 0 Then
      extraiRep = Left(s, pos) & extraiRep(Mid(s, pos + 1))
    Else
      extraiRep = ""
    End If

End Function

Public Function nomCourt(ByVal s As String, Optional ByVal SansExtension As Boolean = True) As String
'Extrait le nom principal d'un fichier (sans son chemin, éventt sans son extension)
'NB FileSystemObject : GetBaseName retourne cette valeur (sans chercher à vérifier la validité ni l'existence du chemin)
  If SansExtension Then
    nomCourt = gtFso.GetBaseName(s)
  Else
    nomCourt = gtFso.GetFileName(s)
  End If
  
End Function

Public Function Extension(s As String, Optional Conversion As VbStrConv) As String
' Retourne l'extension d'un nom de fichier (éventuellement en majuscules ou minuscules)
  
  Extension = gtFso.GetExtensionName(s)
  If Conversion <> 0 Then Extension = StrConv(Extension, Conversion)
    
End Function


Public Function supprSlash(ByVal Texte As String) As String
  If Right(Texte, 1) = "\" And Len(Texte) > 3 Then
    supprSlash = Left(Texte, Len(Texte) - 1)
  Else
    supprSlash = Texte
  End If
End Function


Public Function substPtDecimalRegional(ByVal s As String, Optional ByVal Regional As Boolean) As String
' fonction appelée pour remplacer le point décimal par une virgule ou réciproquement
' ceci permet aux fontions Cdbl et IsNumeric (en particulier) de fonctionner correctement
' Enfin, le drapeau DXFBinaire provient du fait que la lecture par Get,suivi de la fonction CStr peut venir remettre une virgule dans le nombre

  Dim pos%
  Dim vraipoint As String * 1, fauxpoint As String * 1
  
    
  If Regional Then
    vraipoint = Chr(gtPtDecimal)
    fauxpoint = "."
  Else
    vraipoint = "."
    fauxpoint = Chr(gtPtDecimal)
  End If
    
  If vraipoint <> fauxpoint Then
    pos = InStr(s, fauxpoint)
    While pos <> 0
      Mid(s, pos) = vraipoint
      pos = InStr(s, fauxpoint)
    Wend
  End If
  
  substPtDecimalRegional = s
  
End Function

Public Sub iniTextBox(ByVal txtBox As VB.TextBox)
' Mise en surbrillance d'une Zone TextBox sur la longueur effective du texte
On Error Resume Next

  With txtBox
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub


Public Function controleNumeric(ByVal txtBox As VB.TextBox, ByVal StrictPositif As Boolean) As Boolean
' Controle qu'une donnée est bien numérique, et éventuellement strictement positive
Dim Valeur As String

  Valeur = txtBox.Text
  
  If Valeur = "" Then
    controleNumeric = False
  ElseIf Not IsNumeric(Valeur) Then
    MsgBox IDm_Numerique, vbYes + vbExclamation
    controleNumeric = True
  ElseIf StrictPositif And CDbl(Valeur) <= 0 Then
    MsgBox IDm_SuperieurAZero, vbYes + vbExclamation
    controleNumeric = True
  Else
    controleNumeric = False
  End If

End Function

Public Sub gtKey(ByRef Touche As Integer, Optional ByVal Shift As Integer, Optional ByVal Down As Boolean)
  If Shift = 0 And Down Then
    If Touche = vbKeyDecimal Then alertVirgule = True
  ElseIf alertVirgule Then
    Touche = gtPtDecimal
    alertVirgule = False
  End If
End Sub

Public Sub ActivMenu(ByVal Activation As Boolean, Optional ByVal Feuille As Form)
Dim i%

' Dans une future version, l'instruction suivante devrait suffire
' Il faudra gérer autrement le pas à pas
  MDIGiration.Enabled = Activation

End Sub

Public Function gtInteractionEnCours() As Boolean
  If Not fCourante Is Nothing Then
    gtInteractionEnCours = fCourante.flagZoom Or (fCourante.flagPAN And gtOutilZoom = TOOL_PAN) Or fCourante.flagOrigineLibre Or fCourante.flagOrigineGuidée
  End If
End Function

Public Sub DefautCurseur()

  If ObjetDessin Is Nothing Then Exit Sub
  
  If gtCoordActif Or gtOutil = TOOL_AUCUN Then
    ObjetDessin.MousePointer = vbDefault
  Else
    ObjetDessin.MousePointer = vbCrosshair
    If (gtOutil = TOOL_DIST Or gtOutil = TOOL_ANGLEDYN) Then
      DoEvents
      'Modif v3.3(25/04/07) : On évite ainsi de charger frmCadrage inutilement
      'frmCadrage.cmdDésigner(1) = True
      PréparerZoom 1
    End If
  End If

End Sub

Public Sub PréparerZoom(ByVal Index As Integer, Optional ByVal ContexteCadrage As Boolean)

  If Index = 1 Then ' PAN ou outil de mesure
    If FichierJournal Then Write #numFichLog, "PAN"
    fCourante.flagPAN = True
  Else  ' Zoom
    If FichierJournal Then Write #numFichLog, "ZOOM fenêtre"
    fCourante.flagZoom = True
    bgTexte ID_ClicContextuel
  End If
  
  With ObjetDessin
    .SetFocus
    .MousePointer = vbCustom
    If gtOutilZoom = TOOL_PAN Then
      .MouseIcon = MDIGiration.ilsGiration.ListImages("curPAN").Picture
       bgTexte ID_Premier & " " & ID_Point & " : "
      MDIGiration.staMesure.Panels("Distance").Text = ""
    ElseIf gtOutilZoom = TOOL_ZOOM Then
      .MouseIcon = MDIGiration.ilsGiration.ListImages("curZoom").Picture
      MDIGiration.staMesure.Panels("Distance").Text = ""
    ElseIf gtOutil = TOOL_DIST Then
      .MouseIcon = MDIGiration.ilsGiration.ListImages("curDistance2").Picture
    ElseIf gtOutil = TOOL_ANGLEDYN Then
      .MouseIcon = MDIGiration.ilsGiration.ListImages("curAngle2").Picture
    End If
  End With
  
  If gtOutilZoom <> TOOL_SANSZOOM Then    ' Sinon : Outil de mesure et non pas Zoom
  
    If fCourante.Grille.Visible Then
      fCourante.Grille.Enabled = False
      fCourante.picBoutons.Enabled = False
    Else
      fCourante.fraOrigine.Enabled = False
    End If
  End If

End Sub

Public Sub ExecuteZAvant()
' ZOOM Avant
Dim echel As New PT

  If FichierJournal Then Write #numFichLog, "Vue précédente"
  
  verifRetaillage
  With fCourante
    .Milieux.Remove .Milieux.Count
    .Echelles.Remove .Echelles.Count
'    .Form_Activate : remplacé par les 3 lignes suivantes (AV - 08.02.2000)
    ' Alimentation des variables (déclarées dans GraphStand)  utiles aux fonctions de mise à l'échelle
    Set gtMil = .Milieux(.Milieux.Count)
    Set echel = .Echelles(.Echelles.Count)            ' echel est une variable locale, pour déterminer gtFacteurZoom
    gtFacteurZoom = Min(echel.X, echel.Y)
    ParamEcranZoom
  End With
  
  TerminerRecadrage
  
End Sub

Public Sub ExecuteZArrière()
Dim basGauche As New PT, hautDroit As New PT
Dim deltaX As Single, deltaY As Single

  basGauche.X = fCourante.pmin.X
  basGauche.Y = fCourante.pmin.Y
  hautDroit.X = fCourante.pmax.X
  hautDroit.Y = fCourante.pmax.Y
  deltaX = Arrond(hautDroit.X - basGauche.X, 2) / 2
  deltaY = Arrond(hautDroit.Y - basGauche.Y, 2) / 2

  ExecuteZoom basGauche.X - deltaX, basGauche.Y - deltaY, hautDroit.X + deltaX, hautDroit.Y + deltaY
  
End Sub

'***************************************************************************************
' Exécution du Zoom suite à saisie de l'opérateur
'***************************************************************************************
Public Sub ExecuteZoom(X As Single, Y As Single, x0 As Single, y0 As Single)
Dim basGauche As New PT, hautDroit As New PT
Dim deltaX As Single, deltaY As Single

  ObjetDessin.MousePointer = vbDefault '0
  
  If gtOutilZoom = TOOL_ZOOM Then
    basGauche.X = Arrond(trEchelX(X, True), 2)
    basGauche.Y = Arrond(trEchelY(Y, True), 2)
    hautDroit.X = Arrond(trEchelX(x0, True), 2)
    hautDroit.Y = Arrond(trEchelY(y0, True), 2)
    
  ElseIf gtOutilZoom = TOOL_ZARRIERE Then
    basGauche.X = X
    basGauche.Y = Y
    hautDroit.X = x0
    hautDroit.Y = y0
    
  Else  ' PAN ou Outil de mesure
    deltaX = Arrond(trEchelX(x0, True) - trEchelX(X, True), 2)
    deltaY = Arrond(trEchelY(y0, True) - trEchelY(Y, True), 2)
    
    ' Outils de mesure : Affichage du résultat et relance de la commande
    If gtDistance Or gtAngleDyn Then
      If gtDistance Then
        MDIGiration.staMesure.Panels("Distance").Text = "Distance : " & Format(Sqr(Carre(deltaX) + Carre(deltaY)), "#0.0#")
      End If
      If fCourante.fraOrigine.Visible Then fCourante.fraOrigine.Enabled = True
      gtDistance = False: gtAngleDyn = False
      bgTexte ""
      
      'cmdDésigner(1) = True
      ' ligne ci-dessus remplacée par la suivante
      PréparerZoom 1
      Exit Sub
    
    Else  'Commande PAN normale
      basGauche.X = fCourante.pmin.X - deltaX
      basGauche.Y = fCourante.pmin.Y - deltaY
      hautDroit.X = fCourante.pmax.X - deltaX
      hautDroit.Y = fCourante.pmax.Y - deltaY
    End If
  End If
    
' détection d'un plantage possible, si le Zoom amène à une division par zéro
  If hautDroit.X = basGauche.X Or hautDroit.Y = basGauche.Y Then
    MsgBox IDm_SeuilZoom, vbExclamation
    Rafraichir
  
  Else
    verifRetaillage
    CalcEchelle bg:=basGauche, hd:=hautDroit
    TerminerRecadrage
  End If
       
End Sub

Public Sub TerminerRecadrage()
  
  gtOutilZoom = TOOL_SANSZOOM

  With fCourante
    .cmdZAvant.Enabled = (.Milieux.Count > 1)
    MDIGiration.tbrGiration.Buttons("btnZAvant").Enabled = (.Milieux.Count > 1)
    MDIGiration.mnuZoom(TOOL_ZAVANT) = (.Milieux.Count > 1)

    If Not EstChargée(frmPas) Then
' Ajout AV 08.02.2000 : Le Cadrage peut être appelé depuis le pas à pas
      .picBoutons.Enabled = True
    End If
  End With
  
  If fCourante.Grille.Visible Then
    fCourante.Grille.Enabled = True
  Else
    fCourante.fraOrigine.Enabled = True
  End If

  Rafraichir
  
  DefautCurseur

End Sub

Public Sub CalcDuree(ByVal premier As Boolean)
  Dim heure As Date
  Static DEBUT As Double
  Dim fin As Double, duree As Double
  Dim nbheure As Integer, nbminute As Integer, nbsecond As Integer

        heure = Time
        If premier Then
          DEBUT = 3600 * Hour(heure) + 60 * Minute(heure) + Second(heure)
        Else
          fin = 3600 * Hour(heure) + 60 * Minute(heure) + Second(heure)
          duree = fin - DEBUT
         MsgBox "Durée d'exécution : " & CStr(duree)
        End If

End Sub

Public Sub Resol(ByVal ThisForm As Form, ByRef Width As Single, ByRef Height As Single)
Dim DesignX%, DesignY%, FacteurX, FacteurY As Single, i%, j%
Dim nbligneGrille As Integer, lgTexte As Variant
Const COEFFLAG = 0.7
Dim coefg As Double
Dim sommeWidth As Double ', diffWidth As Double
Dim controle As VB.Control

  DesignX = 12      ' sur le portable IBM par ex. TwipsPerPixelX=15
  DesignY = 12
  FacteurX = DesignX / Screen.TwipsPerPixelX
  FacteurY = DesignY / Screen.TwipsPerPixelY
'  FacteurX = 12 / 15
'  FacteurX = 1
  
  If FacteurY = 1 Then Exit Sub
  
  If ThisForm.name = "frmTraject" Then
    coefg = ThisForm.Grille.Width / MDIGiration.Width
  Else
    coefg = ThisForm.fraCarGeom.Width / MDIGiration.Width
    If coefg < COEFFLAG Then Exit Sub
  End If
    If coefg < COEFFLAG Then
      FacteurX = COEFFLAG / coefg
      FacteurY = FacteurX
  End If
  
  On Error GoTo GestErr
  With ThisForm
    Width = Width * FacteurX
    Height = Height * FacteurY
'    .Move 0, 0, .Width * FacteurX, .Height * FacteurY
    For Each controle In .Controls
      If TypeOf controle Is CommonDialog Or TypeOf controle Is Timer Then
      ElseIf TypeOf controle Is ComboBox Then
        With controle
          If .Style <> 1 Then .Move .Left * FacteurX, .Top * FacteurY, .Width * FacteurX
        End With
      Else
        If TypeOf controle Is vaSpread Then
          With controle
            .UnitType = 2
            For j = 0 To 5
              .ColWidth(j) = .ColWidth(j) * FacteurX
            Next
          End With
        End If
              
      'Modif AV 3.3 (13/11/06) : Réactivation du redimensionnement des controles pour frmVéhicule seulement
      
      '------- v3.2----------
'        If TypeOf controle Is VB.LINE Or TypeOf controle Is VB.Shape Then
'        ElseIf controle.name = "fraOrigine" Then
'
'        Else
'          With controle
'            '.Move .Left * FacteurX, .Top * FacteurY, .Width * FacteurX, .Height * FacteurY
'          End With
'        End If
      
      '------- v3.3----------
        If ThisForm.name = "frmVéhicule" Then
          With controle
            .Move .Left * FacteurX, .Top * FacteurY, .Width * FacteurX, .Height * FacteurY
          End With
        End If
      End If
      '-----fin v3.3----------
      
          
      If TypeOf controle Is vaSpread Then
        With .Grille
          
'          .UnitType = 2 ' passage en Twips
          ' 300 pour hauteur approximative de  la ligne d'en-tête
            nbligneGrille = (.Height - .RowHeight(0)) \ (.RowHeight(1))
          
          If (.Height - .RowHeight(0)) Mod (.RowHeight(1)) > 100 Then
            nbligneGrille = nbligneGrille + 1
          End If
            ' Il faut rajouter un petit coefficient pour les interlignes
          .Height = (nbligneGrille) * .RowHeight(1) + (1 + 0.072 * nbligneGrille) * .RowHeight(0)
        sommeWidth = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5)
        
        If coefg < COEFFLAG Then
          .Width = sommeWidth + 0.6 * .ColWidth(0)
        Else
          .Width = sommeWidth + 0.95 * .ColWidth(0)
        End If
          
          .Row = 0
          ThisForm.FontBold = True
          
          lgTexte = ThisForm.TextWidth(ID_Deplacement)
          .Col = 1
          If lgTexte * 1.05 >= .ColWidth(1) Then
            If ThisForm.TextWidth(ID_DeplacementCourt) >= .ColWidth(1) Then
              .Text = Left(ID_Deplacement, 1)
            Else
              .Text = ID_DeplacementCourt
            End If
          Else
            .Text = ID_Deplacement
          End If
          
          lgTexte = ThisForm.TextWidth(ID_Rayongir)
          .Col = 2
          If lgTexte * 1.05 >= .ColWidth(2) Then
            If ThisForm.TextWidth(ID_RayongirCourt) >= .ColWidth(2) Then
              .Text = Left(ID_Rayongir, 1)
            Else
              .Text = ID_RayongirCourt
            End If
          Else
            .Text = ID_Rayongir
          End If
          
          lgTexte = ThisForm.TextWidth(ID_Anglegir)
          .Col = 3
          If lgTexte * 1.05 >= .ColWidth(3) Then
            If ThisForm.TextWidth(ID_AnglegirCourt) >= .ColWidth(3) Then
              .Text = Chr(223)
            Else
              .Text = ID_AnglegirCourt
            End If
          Else
            .Text = ID_Anglegir
          End If
          
          lgTexte = ThisForm.TextWidth(ID_Longueur)
          .Col = 4
          If lgTexte * 1.05 >= .ColWidth(4) Then
'            If ThisForm.TextWidth("Lg") >= .ColWidth(4) Then
              .Text = Left(ID_Longueur, 1)
'            Else
'              .Text = "Lg"
'            End If
          Else
            .Text = ID_Longueur
          End If
          
          lgTexte = ThisForm.TextWidth(ID_VitBraq)
          .Col = 5
          If lgTexte * 1.01 >= .ColWidth(5) Then
            If ThisForm.TextWidth(ID_VitBraqCourt) >= .ColWidth(5) Then
              If ThisForm.TextWidth(ID_VitBraqTresCourt) >= .ColWidth(1) Then
                .Text = "a"
              Else
                .Text = ID_VitBraqTresCourt
              End If
            Else
              .Text = ID_VitBraqCourt
            End If
          Else
            .Text = ID_VitBraq
          End If
        End With
      End If
      
      ' Sur certains postes, la version 32 bits peut planter sur l'affectation de Font.Size ?????? (AV 7/1/98): l'instruction qui suit permet de l'ignorer
      On Error Resume Next
      If TypeOf controle Is TextBox Or TypeOf controle Is label Or TypeOf controle Is CommandButton Or _
          TypeOf controle Is CheckBox Or TypeOf controle Is ComboBox Then 'Or TypeOf controle Is DBGrid Then
'        controle.Font.Size = controle.Font.Size * FacteurX
      End If
     
      ThisForm.MousePointer = vbDefault
    Next controle
    
    If .name = "frmTraject" Then
      .picBoutons.Move .Grille.Left + .Grille.Width + 200
      .fraOrigine.Move .Grille.Left + (.Grille.Width - .fraOrigine.Width) / 2
    End If
  End With  ' ThisForm
  
  Exit Sub
  
GestErr:
  If Err = 384 And ThisForm.WindowState = vbMaximized Then
    Resume Next
  Else
    ErrGeneral
  End If
End Sub

Public Sub SetDeviceIndependentWindow(ByVal ThisForm As Form)
Dim CoefWidth As Single, CoefHeight As Single
Dim Width As Single, Height As Single, Top As Single, Left As Single

  With ThisForm
    Width = .Width
    Height = .Height
    Resol ThisForm, Width, Height
   
   CoefWidth = 0.99   ' 1280*1024
   CoefHeight = 0.921
   If Abs(Screen.Height - 13000) < 50 Then  ' 1152*864
    CoefWidth = 0.989
    CoefHeight = 0.906
  ElseIf Abs(Screen.Height - 11500) < 50 Then ' 1024*768
    CoefWidth = 0.987
    CoefHeight = 0.894
   ElseIf Abs(Screen.Height - 9000) < 50 Then ' Screen.Height=9000 en 800x600
    CoefWidth = 0.984
    CoefHeight = 0.864
   ElseIf Abs(Screen.Height) < 8000 Then ' 640x480
    CoefWidth = 0.98
    CoefHeight = 0.828
   End If
   
    On Error GoTo GestErr
    ' Définit la largeur de la feuille.
    Width = MDIGiration.Width * 0.984   ' au lieu de 0,97
    'Width = MDIGiration.Width * 0.97   ' au lieu de 0,97
    Width = MDIGiration.Width * CoefWidth
    If .name = "frmVéhicule" Then
    ' Si le redimensionnement a été effectué, on risque d'avoir une grande marge à droite sur les grands écrans
      Width = Min(.Width, .Illustration(0).Left + .Illustration(0).Width + .fraCarGeom.Left * 2)
    Else        ' 6/11/97
      '.Move 20, 20   ' 6/11/97
      Left = 20
      Top = 20
    End If
    
    If .name = "frmTraject" Then
    ' Définit la hauteur de la feuille. au maxi possible
      Height = MDIGiration.Height * 0.911
      Height = MDIGiration.Height * 0.864    ' Pour la barre d'outils
      Height = MDIGiration.Height * CoefHeight
'      .Height = MDIGiration.Height * 0.895    ' 6/11/97
    ' Positionne la feuille en haut de la feuille MDI
'    Top = MDIGiration.Top       '  commentaire = 6/11/97
    Else
    ' Centre la feuille verticalement.
      Top = (MDIGiration.Height - .Height) / 2
    ' Centre la feuille horizontalement.
      Left = (MDIGiration.Width - Width) / 2   '  6/11/97
    End If
    
    ' Centre la feuille horizontalement.
    '.Left = (MDIGiration.Width - .Width) / 2  ' Commentaire 6/11/97
   
   .Move Left, Top, Width, Height
   
  End With
  
  Exit Sub
  
GestErr:
  If Err = 384 And ThisForm.WindowState = vbMaximized Then
    Resume Next
  Else
    ErrGeneral
  End If
End Sub

'**************************************************
' Sous Win32 :  les boutons ne sont plus 'graissés'
'***************************************************
Public Sub DégraisserFonte(ByVal Feuille As Form)
  Dim controle As VB.Control
  
  For Each controle In Feuille.Controls
    If TypeOf controle Is VB.CommandButton Then
      controle.Font.Bold = False
    End If
  Next
End Sub

'******************************************************************
' Détecte si la bibliothèque est en lecture seule ou en mise à jour
'******************************************************************
Public Function BibliEnMiseAJour(Optional ByVal premier As Boolean) As Boolean
Dim DrapeauExist As Boolean

  DrapeauExist = ExistFich(GirationVeh)
  If FichierProtégé(GirationVeh, MsgLectureSeule:=Not premier, Titre:="Bibliothèque de véhicules") Then
    gtBibliVerrouillée = True
    BibliEnMiseAJour = True
  ElseIf Not DrapeauExist Then
    ' Effacement du fichier créé à vide par le test de protection
    Kill GirationVeh
  End If
  
End Function

'******************************************************************
' Détecte si le fichier est en lecture seule ou en mise à jour
'******************************************************************
Public Function FichierProtégé(ByVal NomFich As String, Optional ByVal MsgLectureSeule As Boolean = True, Optional ByVal Titre As String, Optional ByVal LectureSeuleAutorisée As Boolean) As Boolean
Dim numFich As Integer
Dim f As Scripting.File
Dim lblFichier As String

  lblFichier = "Fichier " & NomFich
  
  If NomFich = "" Then Exit Function
  If Not ExistFich(NomFich) Then Exit Function

On Error GoTo GestErr

  numFich = FreeFile
  Open NomFich For Append Lock Write As numFich
  Close numFich

    ' Détection d'une protection système en écriture
  Set f = gtFso.GetFile(NomFich)
  If (f.Attributes And ReadOnly) = ReadOnly Then Err.Raise 75
    ' Détection d'un verrouillage en écriture par un autre utilisateur
  numFich = FreeFile
  Open NomFich For Random Lock Write As numFich
  Close numFich
  
  If LectureSeuleAutorisée Then FichierProtégé = False
  
Exit Function

GestErr:
  FichierProtégé = True
  If Err = 75 Then  ' ReadOnly ou Append Lock Write
    If MsgLectureSeule Then MsgBox lblFichier & " en lecture seule", vbExclamation, Titre
    If LectureSeuleAutorisée Then
      MsgLectureSeule = False          ' Pour ne pas avoir le message 2 fois
      Resume Next
    End If
  ElseIf Err = 55 Then  ' Append Lock Write
    MsgBox lblFichier & " déjà ouvert" & vbCrLf & "Enregistrez le d'abord sous un autre nom", vbExclamation, NomFich
  ElseIf Err = 70 Then  ' Append Lock Write ou Random Lock Write
    MsgBox lblFichier & " en cours d'utilisation", vbExclamation, Titre
  Else
    ErrGeneral "Girstand : FichierProtégé"
  End If

End Function

'******************************************************************************
' Message d'erreur non fatale non gérée par Amyos
'******************************************************************************
Public Sub ErreurNonFatale(Optional ByVal MsgErreur As String)
Dim MsgEntete As String

  MsgEntete = "Erreur : " & CStr(Err) & vbCrLf & CStr(Err.Description)
  
  If Len(MsgErreur) > 0 Then
    MsgErreur = MsgEntete & vbCrLf & vbCrLf & "Fonction : " & MsgErreur
  Else
    MsgErreur = MsgEntete
  End If
  
  MsgBox MsgErreur, vbExclamation + vbSystemModal, "Anomalie Giration"
  
End Sub


