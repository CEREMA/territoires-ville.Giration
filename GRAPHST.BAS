Attribute VB_Name = "GraphStand"
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module standard : GraphStand   -   GRAPHST.BAS
'   Module comportant toutes les fonctions graphiques

'**************************************************************************************

Option Explicit

Public tCouleur(255) As Integer, InvTCouleur(15) As Integer
Public pi As Double
'Public gtPlans   As Collection
Public gtPlans As Plans
Public gtEffacement As Boolean
Public flagVéhicule As Boolean, flagEcranZoom As Boolean
Public cptcar%, passage%
Public xMaxPicture As Double, yMaxPicture As Double
Public gtOrx As Single, gtOry As Single
Public gtMil As New PT, gtFacteurZoom As Double
Public Wpmin As New PT, Wpmax As New PT


Public Sub tableCouleur()
  Dim i As Integer
  ' transformation des couleurs AutoCAD en couleur VB (pour la fonction QBColor)
  ' utilisée par les différentes fonctions 'dessin'
  
  tCouleur(0) = 0    ' noir
  tCouleur(1) = 4    ' rouge
  tCouleur(2) = 6    ' jaune
  tCouleur(3) = 2    ' vert
  tCouleur(4) = 3    ' cyan
  tCouleur(5) = 1    ' bleu
  tCouleur(6) = 5    ' magenta
  tCouleur(7) = 0    ' blanc transformé en noir
  tCouleur(8) = 8    ' gris
  For i = 1 To 7
    tCouleur(i + 8) = tCouleur(i) + 8
    tCouleur(i) = tCouleur(i + 8)         ' couleurs + vives systématiqut
    ' Ajout AV 28.03.2000 pour la v3.0.2 : on essaye de restituer qq couleurs spéciales pour le retour d'export
    If i = 3 Then
      tCouleur(i + 8) = 13
    ElseIf i = 5 Then
      tCouleur(i + 8) = 5
    ElseIf i = 7 Then
      tCouleur(i + 8) = 4
    End If
  Next
  
  For i = 16 To 255
    tCouleur(i) = 0  ' noir
  Next i
    
End Sub

Public Sub InvTableCouleur()
Dim i As Integer
  ' transformation des couleurs  VB (pour la fonction QBColor) en couleur AutoCAD
  ' utilisée par les différentes fonctions 'dessin'
  
  InvTCouleur(0) = 7    ' blanc ou noir (selon Installation d'AutoCAD)
  InvTCouleur(15) = 7    ' blanc ou noir (selon Installation d'AutoCAD)
  InvTCouleur(7) = 9    ' gris clair
  InvTCouleur(8) = 8    ' gris
  InvTCouleur(1) = 5    ' bleu
  InvTCouleur(9) = 5    ' bleu
  InvTCouleur(2) = 3    ' vert
  InvTCouleur(10) = 3    ' vert
  InvTCouleur(3) = 4    ' cyan
  InvTCouleur(11) = 4    ' cyan
  InvTCouleur(12) = 1    ' rouge
  InvTCouleur(13) = 11    ' rose
  InvTCouleur(6) = 2    ' jaune
  InvTCouleur(14) = 2    ' jaune
  InvTCouleur(4) = 15    ' marron
  InvTCouleur(5) = 13     ' magenta/mauve
  
End Sub

Public Function rechCouleur(ByVal couleur As Long, ByVal nomPlan As String, ByVal Entité As INSERT) As Long

  If pourEchelle Then rechCouleur = couleur: Exit Function

' Couleur des entités insérées dans un bloc:  cf Manuel utilisateur AutoCAD 12 ch.10 p 394 (Blocs et plans)
  If Not Entité Is Nothing Then              ' Insert
    If couleur = 0 Then       'couleur DUBLOC
      rechCouleur = rechCouleur(Entité.couleur, Entité.nomPlan, Nothing)
    Else
      If nomPlan = "0" Then    ' plan caméléon: couleur du plan du bloc
          rechCouleur = rechCouleur(couleur, Entité.nomPlan, Nothing)
      Else
        rechCouleur = rechCouleur(couleur, nomPlan, Nothing)
      End If
    End If
    
  Else                                                                'Entité simple
    If couleur = 256 Then   'couleur DUPLAN
      rechCouleur = gtPlans(nomPlan).couleur
    Else
      rechCouleur = couleur
    End If
  End If

End Function

Public Sub ptEchelle(ByVal p As PT, ByVal Coul As Long)
  ' Dessin d 'un point avec mise à l'échelle et calage de l'origine
Dim xa, ya As Double
      
  If frmAnnulImpr.Annul Then
    Err.Raise 102
  End If
  
'      If pourEchelle Then minimax p.X, p.Y: Exit Sub
  
  
  xa = trEchelX(p.X, False)
  ya = trEchelY(p.Y, False)
  
  If xa < 0 Or ya < 0 Or xa > xMaxPicture Or ya > yMaxPicture Then Exit Sub
  
  On Error GoTo GestErr
  
  If gtEffacement Then
    ObjetDessin.PSet (xa, ya), ObjetDessin.BackColor
  Else
    ObjetDessin.PSet (xa, ya), Coul
  End If

Exit Sub
      
GestErr:
  If Err <> 6 And Err <> 5 Then
    MsgBox IDm_AnomalieDessin & CStr(Err)
  End If
      
End Sub

Public Sub desEchelle(ByVal pa As PT, ByVal pb As PT, ByVal Coul As Long)
  ' Dessin d 'un segment avec mise à l'échelle et calage de l'origine
Dim p1 As New PT, p2 As New PT
     
  p1.X = trEchelX(pa.X, False)
  p1.Y = trEchelY(pa.Y, False)
  p2.X = trEchelX(pb.X, False)
  p2.Y = trEchelY(pb.Y, False)
  
  If Not clip(p1, p2) Then
  
On Error GoTo GestErr
  
    If gtEffacement Then
      ObjetDessin.Line (p1.X, p1.Y)-(p2.X, p2.Y), ObjetDessin.BackColor
  '        Write #numFichLog, "efface"
    ElseIf frmAnnulImpr.Annul Then
      Err.Raise 102
    Else
      ObjetDessin.Line (p1.X, p1.Y)-(p2.X, p2.Y), Coul
  '        Write #numFichLog, xa, ya, xb, yb
      If passage <> 0 Then pasSuivant
    End If
  End If
  
  Set p1 = Nothing
  Set p2 = Nothing

Exit Sub
      
GestErr:
  If Err <> 6 And Err <> 5 Then
    MsgBox IDm_AnomalieDessin & CStr(Err)
  End If
Exit Sub

End Sub

Public Sub arcEchelle(ByVal p As PT, ByVal R As Single, ByVal Coul As Long, _
                    Optional ByVal Angle1 As Double, Optional ByVal Angle2 As Double = 360)
  ' dessin d'un arc de cercle avec mise à l'échelle et calage de l'origine
Dim xa As Double, ya As Double
Dim AngleDep As Double
Dim CerclePur As Boolean
Dim Rayon As Single

  If frmAnnulImpr.Annul Then
    Err.Raise 102
  End If
      
  CerclePur = (Angle1 = 0# And Angle2 = 360#)
      
' Modif AV 28/04/98 : Sur certains fichiers DXF un peu tordus (polylignes avec nombreux petits arcs
' VB calcule mal les arcs --> on les transforme en segments de ligne (deltaAngle < 1°)
  If Not CerclePur Then
    If Abs(Angle1 - Angle2) < 1# Then
      Dim Alpha As Double, Beta As Double
      Dim pa As New PT, pb As New PT
      DoEvents
      Alpha = angConv(Angle1, radian)
      Beta = angConv(Angle2, radian)
      pa.X = p.X + R * Cos(Alpha)
      pa.Y = p.Y + R * Sin(Alpha)
      pb.X = p.X + R * Cos(Beta)
      pb.Y = p.Y + R * Sin(Beta)
      desEchelle pa, pb, Coul
      Exit Sub
    End If
  End If
        
  xa = trEchelX(p.X, False)
  ya = trEchelY(p.Y, False)
  Rayon = R * gtFacteurZoom
        
  On Error GoTo GestErr
  
  If gtEffacement Then Coul = fCourante.picTrajectoire.BackColor
  
  If CerclePur Then      ' cercle
      ObjetDessin.Circle (xa, ya), Rayon, Coul
      If passage <> 0 Then pasSuivant
  
  Else                                          ' arc
    ObjetDessin.Circle (xa, ya), Rayon, Coul, angConv(Angle1, radian), angConv(Angle2, radian)
  End If

Exit Sub
      
GestErr:
  If Err <> 6 And Err <> 5 Then
    MsgBox IDm_AnomalieDessin & CStr(Err)
  End If
Exit Sub

End Sub

Public Sub Rafraichir()
    
 On Error GoTo GestErr
 
    If FichierJournal Then Write #numFichLog, "Rafraichir"
    
    ActivMenu Activation:=False
    
    verifRetaillage
    
    If Not ObjetDessin Is Printer Then
      ObjetDessin.Cls
      Graduer
    Else
      If frmAnnulImpr.Annul Then
        ActivMenu Activation:=True
        Err.Raise 102
      End If
      Printer.FillStyle = vbHorizontalLine
      Printer.Print ""
      Printer.FillStyle = vbFSTransparent
    End If
       
    desFDP
    
    If EstChargée(frmPas) Then
    ' Rafraichir a été appelé depuis le Pas à Pas (demande de Zoom) : On ne redessine qu'une position du véhicule
      ActivMenu Activation:=True
      frmPas.dessiner RetourZoom:=True
      Exit Sub
    End If

    If frmAnnulImpr.Annul Then
        ActivMenu Activation:=True
        Err.Raise 102
    End If
    
    sensTracteur = OptionAffichage(NumFille).sens
    cptcar = 0
    passage = 1
    fCourante.dessiner "TOUT"
    passage = 0
    sensTracteur = False
    fCourante.colPlans(GirPlanReserve(7)).couleur = fCourante.picTrajectoire.BackColor

  ' Zoom pendant l'opération de modif origine
  If fCourante.fraOrigine.Visible Then fCourante.RedefOrigine

    ActivMenu Activation:=True
   
  Exit Sub
 
GestErr:
  If Err = 102 Then   ' Annulation de l'impression ou du rafraichissement
    sensTracteur = False
    fCourante.colPlans(GirPlanReserve(7)).couleur = fCourante.picTrajectoire.BackColor
    fCourante.MousePointer = vbDefault
    ActivMenu Activation:=True
    Err.Raise 102   ' Fin du traitement de l'annulation dans frmImprime ou frmOptaff
  End If

  End Sub
  
Public Sub desFDP()
' Dessin du fond de plan
' Cette fonction est appelée par
'      Rafraichir
'      crefille dans GIRATION.BAS, en cas de lecture d'un fichier .GIR
'       import dans GIRSTAND.BAS

Dim Cpt As Integer
Dim chaine As String
Dim Coul As Long
Dim p As PT, objPoint As POINT, objLine As LINE, objCercle As CERCLE, objArc As ARC
Dim objPolyline As POLYLINE, objInsert As INSERT
Dim objSpline As SPLINE

On Error GoTo GestErr


  fCourante.MousePointer = vbHourglass    ' Sablier
'  CalcDuree True
  passage = 1
  
  If OptionAffichage(NumFille).FDP.ADessiner Then
    Coul = OptionAffichage(NumFille).FDP.couleur
      ' si coul=-1, on respecte les couleurs d'origine du fond de plan
    With fCourante.GraphFDP
      For Each objPoint In .gphPoints
        objPoint.dessiner Coul
      Next
      For Each objLine In .gphLines
        objLine.dessiner Coul
      Next
      For Each objCercle In .gphCercles
        objCercle.dessiner Coul
      Next
      For Each objPolyline In .gphPolylines
        objPolyline.dessiner Coul
      Next
      For Each objSpline In .gphSplines
        objSpline.dessiner Coul
      Next
      For Each objArc In .gphArcs
        objArc.dessiner Coul
      Next
      For Each objInsert In .gphInserts
        objInsert.dessiner Coul
      Next
      
    End With
'    ObjetDessin.DrawWidth = 1
  End If
  
  
  fCourante.MousePointer = vbDefault
 ' CalcDuree False
  passage = 0
  DoEvents
  
Exit Sub
GestErr:
  ErrLocal Procédure:="desFDP"
  Err.Raise Err.Number

End Sub

Public Sub pasSuivant()
  
  passage = passage + 1
  If passage Mod 100 = 0 Then DoEvents
  
End Sub

Public Sub Graduer()
Dim gradmini As Integer, increment As Integer, pas As Integer
Dim LargImage As Single, lg1 As Single, dern As Single
Dim xminImage As Double, xmaxImage As Double, yminImage As Double, ymaxImage As Double
Dim chaine As String, trait As Single
Dim pa As New PT, pb As New PT
Dim Coul As Long
Dim InitFormat As Boolean

On Error GoTo GestErr

  If OptionAffichage(NumFille).Graduations.ADessiner Then
    Coul = OptionAffichage(NumFille).Graduations.couleur
  
    gradmini = 5    ' graduation mini de 5 en 5, sinon de 10 en 10
  
    With fCourante
        ObjetDessin.ForeColor = Coul
        xminImage = .pmin.X
        xmaxImage = .pmax.X
        yminImage = .pmin.Y
        ymaxImage = .pmax.Y
        trait = 50 / gtFacteurZoom
    '    If ObjetDessin Is Printer Then trait = trait / 567
    End With

        
    '= GRADUER l'AXE des X
    
    LargImage = xmaxImage - xminImage
    lg1 = LargImage / (15 * gradmini) + 0.5     ' 15 graduations  sur l'axe des X
    increment = Fix(lg1) * gradmini
    If increment Mod 10 = gradmini Then
      pas = gradmini
    ElseIf increment < 5 Then
      lg1 = LargImage / 15 + 0.5    ' 15 graduations  sur l'axe des X de 1 en 1
      increment = Fix(lg1)
      If increment = 0 Then increment = 1
      pas = 1
    Else
      pas = 10
    End If
    
    dern = Int((xminImage + pas) / pas) * pas
    pa.X = dern
    pa.Y = ymaxImage - trait
    pb.X = dern
    pb.Y = ymaxImage
    Do
      'Dessiner le trait de graduation
      desEchelle pa, pb, Coul
      'Ecrire le texte de gradutaion
      chaine = superFix(dern, 6)
      With ObjetDessin
        If (.Width - .CurrentX) > 50 And .CurrentX > 200 Then
        'Ne pas écrire la graduation si on est trop près du bord gauche ou droit
          If Not InitFormat Then
            chaine = Format(dern, "##0")
            InitFormat = True
          End If
          'Centrer à peu près le texte sur le trait vertical
          .CurrentX = .CurrentX - .TextWidth(chaine) / 2 ' Len(chaine) * 40
          'Ecrire 50 twips sous le bord haut pour se dégager du trait
          .CurrentY = .CurrentY + 50
          ObjetDessin.Print chaine
        End If
      End With
      dern = dern + increment
      pa.X = dern
      pb.X = dern
    Loop While dern < xmaxImage
            
    '= GRADUER l'AXE des Y
    LargImage = ymaxImage - yminImage
    lg1 = LargImage / (10 * gradmini) + 0.5        ' 10 graduations  sur l'axe des Y
    increment = Fix(lg1) * gradmini
    If increment Mod 10 = gradmini Then
      pas = gradmini
    ElseIf increment < 5 Then
      lg1 = LargImage / 10 + 0.5     ' 10 graduations  sur l'axe des Y de 1 en 1
      increment = Fix(lg1)
      If increment = 0 Then increment = 1
      pas = 1
    Else
      pas = 10
    End If
    
    dern = Int((yminImage + pas) / pas) * pas
    pa.X = xminImage
    pa.Y = dern
    pb.X = xminImage + trait
    pb.Y = dern
    InitFormat = False
    Do
      'Dessiner le trait de graduation
      desEchelle pa, pb, Coul
        'Ecrire le texte de gradutaion
      chaine = superFix(dern, 6)
      With ObjetDessin
        If (.Height - .CurrentY) > 50 And .CurrentY > 200 Then
        'Ne pas écrire la graduation si on est trop près du bord haut ou bas
          If Not InitFormat Then
            chaine = Format(dern, "##0")
            InitFormat = True
          End If
          'Ecrire à 20 twips à droite du trait pour se dégager du trait
          .CurrentX = .CurrentX + 20
          'Centrer à peu près verticalement le texte sur le trait horizontal
          .CurrentY = .CurrentY - 90
          ObjetDessin.Print chaine
        End If
      End With
      dern = dern + increment
      pa.Y = dern
      pb.Y = dern
    Loop While dern < ymaxImage
  
  End If
   
Exit Sub

GestErr:
'  Resume
End Sub

Public Sub DesCarross(ByVal numRemorque As Integer)
' Cette fonction est appelée par
'     FinDesCarross dans frmTraject
'     cmdFermer_Click et txtOri_Change en cas de modification de l'origine dans frmTraject
'     creFille si Nouvelle Trajectoire dans GIRATION.BAS

Dim p As New PT

  If OptionAffichage(NumFille).Carross(numRemorque).ADessiner Then
    flagVéhicule = True
    If numRemorque = TRACTEUR Then
    ' On ne dessine le nez du tracteur (plan GirPlanReserve(7)) que si c'est demandé en options d'affichage
    '!!AV -->.NET : MOD
      If sensTracteur And cptcar Mod OptionAffichage(NumFille).pas(1) = 0 Then
        fCourante.colPlans(GirPlanReserve(7)).couleur = OptionAffichage(NumFille).Carross(numRemorque).couleur
      Else
        fCourante.colPlans(GirPlanReserve(7)).couleur = fCourante.picTrajectoire.BackColor
      End If
      cptcar = cptcar + 1
      fCourante.Carross.gphInserts(1).dessiner -2
    Else
      fCourante.Carross.gphInserts(numRemorque + 1).dessiner fCourante.colPlans(GirPlanReserve(numRemorque)).couleur
    End If
    flagVéhicule = False
  End If

End Sub

'**************************************************************
'*Dessin d'un élément de carrosserie pour saisie de l'origine
'**************************************************************
Public Sub DesCarrossElement(ByVal num As Integer, ByVal p As PT, ByVal rot As Double)

Dim pa As PT, pb As PT
Dim i As Integer
Dim Nom As String

On Error GoTo GestErr

  Select Case num
  Case 0  ' Tracteur
    Nom = "Tracteur"
  Case 1  ' Remorque1
    Nom = "Remorque1"
  Case 2  ' Remorque2
    Nom = "Remorque2"
  End Select
  
  For i = 1 To 4
    With fCourante.linCarrosserie(num * 4 + i)
      With fCourante.Carross.gphBlocs(Nom).blkPolylines(1)
        Set pa = TransRot(.Pts(i Mod 4 + 1), p, rot, 1)
      End With
      .X1 = trEchelX(pa.X, toReel:=False)
      .Y1 = trEchelY(pa.Y, toReel:=False)
      fCourante.linCarrosserie(num * 4 + 1 + (i + 2) Mod 4).X2 = .X1
      fCourante.linCarrosserie(num * 4 + 1 + (i + 2) Mod 4).Y2 = .Y1
    End With
  Next
  
  With fCourante.shpPcarac(num)
    .Left = trEchelX(p.X, toReel:=False) - .Width / 2
    .Top = trEchelY(p.Y, toReel:=False) - .Height / 2
  End With

Exit Sub
GestErr:
  Err.Raise Err.Number
  
End Sub


'**************************************************************
'*Dessin de la (des) carrosserie(s) pour saisie de l'origine
'**************************************************************

Public Sub DesCarrossPourOrigine()
Dim objInsert As INSERT
Dim pa As PT, pb As PT, pr As New PT
Dim i As Integer, rot As Double
Dim i1 As Integer, i2 As Integer

On Error GoTo GestErr

Set objInsert = fCourante.Carross.gphInserts(1)
rot = objInsert.rot
Set pr = objInsert.p
DesCarrossElement 0, pr, rot

    
' Nez du tracteur  : 2è polyligne du bloc tracteur
With fCourante.Carross.gphBlocs("Tracteur").blkPolylines(2)
  Set pa = objInsert.TransRotBloc(.Pts(2))
End With

With fCourante
  i1 = (.nbRemorque + 1) * 4 + 1
  i2 = i1 + 1
  .linCarrosserie(i1).X2 = trEchelX(pa.X, toReel:=False)
  .linCarrosserie(i1).Y2 = trEchelY(pa.Y, toReel:=False)
  
  .linCarrosserie(i2).X2 = .linCarrosserie(i1).X2
  .linCarrosserie(i2).Y2 = .linCarrosserie(i1).Y2
  
  .linCarrosserie(i1).X1 = .linCarrosserie(4).X1
  .linCarrosserie(i1).Y1 = .linCarrosserie(4).Y1
  .linCarrosserie(i2).X1 = .linCarrosserie(4).X2
  .linCarrosserie(i2).Y1 = .linCarrosserie(4).Y2
    
  ' Dessin des remorques
  If .nbRemorque > 0 Then
    .Carross.gphInserts(2).rot = rot
    Set pr = .Carross.gphInserts(2).p
    With .Vehic
      pr.X = objInsert.p.X - ((.Dpat(0) + .e(0))) * Cos(angConv(rot, Enradian:=True))
      pr.Y = objInsert.p.Y - ((.Dpat(0) + .e(0))) * Sin(angConv(rot, Enradian:=True))
    End With
    DesCarrossElement 1, pr, rot
    
    If .nbRemorque = 2 Then
      Set objInsert = .Carross.gphInserts(2)
      .Carross.gphInserts(3).rot = rot
      Set pr = .Carross.gphInserts(3).p
      With .Vehic
        pr.X = objInsert.p.X - ((.Dpat(1) + .e(1))) * Cos(angConv(rot, Enradian:=True))
        pr.Y = objInsert.p.Y - ((.Dpat(1) + .e(1))) * Sin(angConv(rot, Enradian:=True))
      End With
      DesCarrossElement 2, pr, rot
    End If
  End If
End With    ' fCourante
  
Exit Sub
GestErr:
  Err.Raise Err.Number
  
End Sub


Public Function trEchelX(ByVal X As Double, toReel As Boolean) As Double
'Transformée de X : d'unités dessin en unités réelles ou réciproquement
' gtOrx est exprimé en unités dessin
' gtMil.X est exprimé en unités réelles
' X est exprimé en unités dessin (toReel=False) ou réeelles (toReel=True)
  If toReel Then
    trEchelX = gtMil.X + (X - gtOrx) / gtFacteurZoom
  Else
    trEchelX = (X - gtMil.X) * gtFacteurZoom + gtOrx
  End If
 
End Function

Public Function trEchelY(ByVal Y As Double, toReel As Boolean) As Double
'Transformée de Y : d'unités dessin en unités réelles ou réciproquement
' gtOry est exprimé en unités dessin
' gtMil.y est exprimé en unités réelles
' Y est exprimé en unités dessin (toReel=False) ou réeelles (toReel=True)
  If toReel Then
    trEchelY = gtMil.Y - (Y - gtOry) / gtFacteurZoom
  Else
    trEchelY = -(Y - gtMil.Y) * gtFacteurZoom + gtOry
  End If
  
End Function

Private Function codeClip(ByVal p As PT) As Integer

'En VB (v4 à v6 tout au moins) True=-1 et False=0
' On remultiplie par -1 chaque valeur pour faire 'ressortir' positivement les bits intéressants
  codeClip = (p.X < 0#) * -8 + (p.X > xMaxPicture) * -4 + (p.Y < 0#) * -2 + (p.Y > yMaxPicture) * -1

End Function

Private Function clip(ByVal p1 As PT, ByVal p2 As PT) As Boolean
Dim c1 As Integer, c2 As Integer  ' codes binaires du clipping pour p1 et p2
Dim dX As Double, dY As Double    ' deltaX et deltaY entre p1 et p2

  c1 = codeClip(p1)
  c2 = codeClip(p2)
  
  Do While c1 Or c2
    ' Un point au moins hors de la vue
    If c1 And c2 Then
      ' Sortir de la boucle : les 2 points sont hors de la Vue
      c1 = 0
      c2 = 0
      clip = True
    Else
    
      dX = p2.X - p1.X
      dY = p2.Y - p1.Y
      
      If c1 Then
        ' p1 est hors de la vue - redéfinir p1 à l'intersection de p1p2  et d'un bord
        c1 = NewCodeClip(p1, c1, dX, dY)
        ' p2 est hors de la vue - redéfinir p2 à l'intersection de p1p2  et d'un bord
      Else
        c2 = NewCodeClip(p2, c2, dX, dY)
      End If
      
    End If
  Loop
  
End Function

Private Function NewCodeClip(ByVal p As PT, ByVal c As Integer, ByVal dX As Double, ByVal dY As Double) As Integer

  If c And 8 Then
    'Tronquage du bord gauche
    p.Y = p.Y + dY * (-p.X) / dX
    p.X = 0#
    
  ElseIf c And 4 Then
    'Tronquage du bord droit
    p.Y = p.Y + dY * (xMaxPicture - p.X) / dX
    p.X = xMaxPicture
    
  ElseIf c And 2 Then
    'Tronquage du bord inférieur
    p.X = p.X + dX * (-p.Y) / dY
    p.Y = 0#

  ElseIf c And 1 Then
    'Tronquage du bord supérieur
    p.X = p.X + dX * (yMaxPicture - p.Y) / dY
    p.Y = yMaxPicture
  End If
  
  NewCodeClip = codeClip(p)

End Function

Public Sub ToutVoir()
' Calcul des extensions minimales et maximales du dessin, en fonction des options d'affichage

' Cette fonction est appelée par
 ' le Zoom Etendu dans frmCadrage
 ' la fonction creFille si lecture d'un fichier existant dans GIRATION.BAS
' la fonction EchelAjustée dans IMPRIME.FRM
' ECRDXF si export

  Wpmin.X = 1000000000000#
  Wpmin.Y = 1000000000000#
  Wpmax.X = -1000000000000#
  Wpmax.Y = -1000000000000#

'      Calcul des limites utiles de la trajectoire
  pourEchelle = True
  fCourante.dessiner "TOUT"
  pourEchelle = False
  
  ' Réincorporation éventuelle du FDP
  CalcEchelle
   
End Sub


Public Sub CalcEchelle(Optional ByVal bg As PT, Optional ByVal hd As PT)
Dim Etendu As Boolean
Dim w As Single, h As Single
Dim demiLargReel As Double, demiLongReel As Double
Dim echel As New PT
Dim pmin As New PT, pmax As New PT
Dim FDPPresent As Boolean

On Error GoTo GestErr

  If bg Is Nothing Then Etendu = True
  
' Recalcul de l'échelle en vue d'un Zoom Etendu
' numechelle permet de se repositionner sur la vue en cours en cas de nouvel import

  If Etendu Then
     FDPPresent = fCourante.pminFDP.X <> 0 And fCourante.pmaxFDP.X <> 0
    If OptionAffichage(NumFille).FDP.ADessiner And FDPPresent Then
      Wpmin.X = Min(fCourante.pminFDP.X, Wpmin.X)
      Wpmin.Y = Min(fCourante.pminFDP.Y, Wpmin.Y)
      Wpmax.X = Max(fCourante.pmaxFDP.X, Wpmax.X)
      Wpmax.Y = Max(fCourante.pmaxFDP.Y, Wpmax.Y)
    End If
  ' L'appel est fait  dans EcrDXF simplement pour déterminer les mini  et maxi
    If flagEcranZoom Then Exit Sub
    
    Set pmin = Wpmin
    Set pmax = Wpmax
    
  Else    ' Zoom fenêtre ou PAN
    Set pmin = bg
    Set pmax = hd
  End If
  

  With fCourante
    With ObjetDessin
      xMaxPicture = .ScaleWidth
      yMaxPicture = .ScaleHeight
      If ObjetDessin Is Printer Then yMaxPicture = yMaxPicture - 2    ' 2cm pour l'en-tête
    End With
    
' 95% pour les marges
      echel.X = xMaxPicture / (pmax.X - pmin.X) * 0.95
      echel.Y = yMaxPicture / (pmax.Y - pmin.Y) * 0.95
      .Echelles.Add echel.X, echel.Y
      gtFacteurZoom = Min(echel.X, echel.Y)
      Set gtMil = pMilieu(pmin, pmax)
      .Milieux.Add gtMil.X, gtMil.Y
     
      ' redéfintion des paramètres réels de l'écran en fonction du Facteur de Zoom
      If Not ObjetDessin Is Printer Then ParamEcranZoom

  End With  ' fCourante

Exit Sub
GestErr:
  Err.Raise 1001
End Sub

Public Sub ParamEcranZoom()
' Calcule les limites frmTraject.picTrajectoire (en  unités réelles) en fonction du Facteur de Zoom

  With fCourante
    .pmin.X = trEchelX(0, True)
    .pmin.Y = trEchelY(yMaxPicture, True)
    .pmax.X = trEchelX(xMaxPicture, True)
    .pmax.Y = trEchelY(0, True)
    .diagonalEcran = Sqr(Carre(.pmax.X - gtMil.X) + Carre(.pmax.Y - gtMil.Y))
    .rayonInscrit = Min(Abs(.pmax.X - gtMil.X), Abs(.pmax.Y - gtMil.Y))
  End With

End Sub

Public Sub OrigineParDefaut()
' A l'initialisation, position l'origine par défaut à 0,0 si
'   pas de Fond de plan ou si
'   le pt 0,0 est visible après un Zoom Etendu fait sur le Fond de plan

  With fCourante
      .IgnoreOrigine = True  ' pour éviter le redessin intempestif du tracteur (évènement txtOri_Change)
      If .pmin.X < 0 And .pmax.X > 0 And .pmin.Y < 0 And .pmax.Y > 0 Then
        .txtOri(0) = 0     'si 0,0 est sur l'image on le prend pour origine par défaut
        .txtOri(1) = 0
      Else              ' sinon, on s'approche du centre de l'image
        .txtOri(0) = Fix(trEchelX(gtOrx, True))
        .txtOri(1) = Fix(trEchelY(gtOry, True))
      End If
      .txtOri(2) = 0 ' direction d'origine=0
      .IgnoreOrigine = False
    
  End With

End Sub

Public Sub verifRetaillage(Optional ByVal Ouverture As Boolean)

  If fCourante.neworx <> 0 Then
  ' Détection d'un redimensionnement de la fenêtre par l'utilisateur
  ' le centre de la vue a été déplacé, mais le Facteur de Zoom reste le même
  ' le flag Ouverture indique l'ouverture d'une nouvelle fenêtre : la collection Milieux n'est encore pas créée
    If FichierJournal Then Write #numFichLog, "VerifRetaillage"
    
    With fCourante
      gtOrx = .neworx
      gtOry = .newory
      .neworx = 0
      .POrigine.X = gtOrx
      .POrigine.Y = gtOry
      If Not Ouverture Then
        Set gtMil = pMilieu(.pmin, .pmax)
        With .Milieux(.Milieux.Count)
          .X = gtMil.X
          .Y = gtMil.Y
        End With
      End If
    End With
  End If

End Sub


