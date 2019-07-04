VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ARC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public nomPlan As String
Public couleur As Long
Public typelign As String
Public p As New PT
'Public Pts As New Pts
Public Rayon As Double, angdep As Double, angfin As Double

'****************************************
' Ecriture de l'objet dans un fichier DXF
'****************************************
Public Sub ecrire()

  ecrArticle "ARC", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
    ecr 100, SansCadrage:=True
    ecr "AcDbCircle"
  End If

  ecrPlan nomPlan
  ecrCouleur couleur
  ecrTypelign typelign
  ecrPoint p
  ecr 40, True
  ecr Rayon
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbArc"
  End If
  
  ecr 50, True
  If OptGen.unite = GRADE Then
    ecr angdep * 0.9
    ecr 51, True
    ecr angfin * 0.9
  Else
    ecr angdep
    ecr 51, True
    ecr angfin
  End If

End Sub

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
Dim i As Integer
Dim xa As Double, ya As Double
Dim pa As New PT, R As Double, a1 As Double, a2 As Double
Dim Coul As Long
  
  If coulObjet = -1 Then
      '  on respecte les couleurs d'origine du fond de plan
    Coul = rechCouleur(couleur, nomPlan, objInsert)
    Coul = QBColor(tCouleur(Abs(Coul)))   ' la couleur peut être négative si plan inactif dans le DXF
  Else
    Coul = coulObjet
  End If
  
  If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
    With objInsert
      Set pa = .TransRotBloc(p)
      R = Rayon * Abs(.echx)
      If .echx < 0 Then
        a1 = 360 - angfin + .rot
        a2 = 360 - angdep + .rot
      Else
        a1 = angdep + .rot
        a2 = angfin + .rot
      End If
      If a1 >= 2 * eqvPI(OptGen.unite) Then a1 = a1 - 2 * eqvPI(OptGen.unite)
      If a2 >= 2 * eqvPI(OptGen.unite) Then a2 = a2 - 2 * eqvPI(OptGen.unite)
    End With
  Else
    Set pa = p
    R = Rayon
    a1 = angdep
    a2 = angfin
  End If
  
  If pourEchelle Then
    minimaxArc pa, R, a1, a2
  Else
    ' dessin
    arcEchelle pa, R, Coul, a1, a2
  End If

End Sub

Private Sub minimaxArc(ByVal p As PT, ByVal R As Double, ByVal Angle1 As Double, ByVal Angle2 As Double)
Dim AngleDep As Double

  If Angle1 < Angle2 Then
    AngleDep = Angle1
  Else
    AngleDep = Angle1 - 360
  End If

  While AngleDep < Angle2
    minimax p.X + R * Cos(AngleDep * pi / 180), p.Y + R * Sin(AngleDep * pi / 180)
    AngleDep = AngleDep + 30
  Wend
  minimax p.X + R * Cos(Angle2 * pi / 180), p.Y + R * Sin(Angle2 * pi / 180)

End Sub
