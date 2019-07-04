VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POLYLINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public nomPlan As String
Public couleur As Long
Public typelign As String
Public ferme As Boolean
Public TypeLissage As Integer
Public bln3D As Boolean
Public Pts As New Pts
Public Arcs As New Arcs

'****************************************
' Ecriture de l'objet dans un fichier DXF
'****************************************
Public Sub ecrire()
Dim p As New PT

  ecrArticle "POLYLINE", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
    ecr 100, SansCadrage:=True
    ecr "AcDb2dPolyline"
  End If

  ecrPlan nomPlan
  ecrCouleur couleur
  ecrTypelign typelign
  
  ecr 66, True    ' Drapeau Sommets suivent"
  ecr 1
  p.x = 0
  p.y = 0
  ecrPoint p        ' le z éventuel indiquerait l'élévation
  If ferme Then
    ecr 70, True
    ecr 1
  End If
  If TypeLissage <> 0 Then    ' TypeLissage =5 ou 6 (B-spline respectivt quadratique ou cubique : cf variable SPLINETYPE d'AutoCAD)
    ecr 75, True
    ecr TypeLissage
  End If
  
  For Each p In Pts
    ecrArticle "VERTEX", AvecID:=VersionDXF > 11
    
    If VersionDXF > 11 Then
      ecr 100, SansCadrage:=True
      ecr "AcDbEntity"
      ecr 100, SansCadrage:=True
      ecr "AcDbVertex"
      ecr 100, SansCadrage:=True
      ecr "AcDb2dVertex"
    End If

    ecrPlan nomPlan
    ecrPoint p
    ecr 42, True
    ecr p.arrondi
    If p.Drapeau <> 0 Then
      ecr 70, True
      ecr p.Drapeau
    End If
  Next
    
  ecrArticle "SEQEND", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
  End If

  ecrPlan nomPlan              ' la doc de le précise pas, mais c'est quand même présent
  
End Sub

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
Dim i As Integer, j As Integer, nbPoints As Integer, nbSegments As Integer, iSuivant As Integer
Dim pa As PT, pb As PT
Dim Coul As Long
Dim mPts As Pts
Dim unArc As ARC

  If coulObjet = -1 Then
      '  on respecte les couleurs d'origine du fond de plan
    Coul = rechCouleur(couleur, nomPlan, objInsert)
    Coul = QBColor(tCouleur(Abs(Coul)))              ' la couleur peut être négative si plan inactif dans le DXF
  ElseIf coulObjet = -2 Then
      ' Cas particulier du dessin de carrosserie (DesCarross)
    Coul = rechCouleur(couleur, nomPlan, objInsert)
  Else
    Coul = coulObjet
  End If
  
  Set mPts = PtsUtiles
  
  nbPoints = mPts.Count
  If ferme Then
    nbSegments = nbPoints
  Else
    nbSegments = nbPoints - 1
  End If
  
  For i = 1 To nbSegments
    iSuivant = i Mod nbPoints + 1
    
    If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
      With objInsert
        Set pa = .TransRotBloc(mPts(i))
        Set pb = .TransRotBloc(mPts(iSuivant))
      End With
    Else
      Set pa = mPts(i)
      Set pb = mPts(iSuivant)
    End If
  
    If pa.arrondi <> 0 Then
      j = j + 1
      Set unArc = Arcs(j)
      unArc.dessiner Coul, objInsert
      
    Else
      If pourEchelle Then
        minimax pa.x, pa.y
      Else
        desEchelle pa, pb, Coul
      End If
    End If
  Next

  If pourEchelle Then minimax pb.x, pb.y

End Sub

Private Sub Class_Initialize()
'Initialiser avec la polyligne avec les valeurs par défaut d'AutoCAD
  couleur = 256         ' DUPLAN
  typelign = "BYLAYER"  ' DUPLAN
  nomPlan = "0"
End Sub

Public Property Get PtsUtiles() As Pts
Dim p As PT
Dim mPts As Pts

  If TypeLissage = 0 Then
    Set PtsUtiles = Pts
  Else
    Set mPts = New Pts
    For Each p In Pts
      With p
        If (.Drapeau And 16) <> 16 Then   ' Si drapeau = 16 : point de controle --> n'est pas sur la courbe
          mPts.Add .x, .y, .arrondi, .Drapeau
        End If
      End With
    Next
    Set PtsUtiles = mPts
    Set mPts = Nothing
  End If

End Property

