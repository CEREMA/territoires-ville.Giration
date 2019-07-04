VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CERCLE"
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
Public Rayon As Double

'****************************************
' Ecriture de l'objet dans un fichier DXF
'****************************************
Public Sub ecrire()

  ecrArticle "CIRCLE", AvecID:=VersionDXF > 11
  
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

End Sub

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
Dim i As Integer
Dim xa As Double, ya As Double
Dim pa As New PT, R As Single
Dim Coul As Long
  
  If coulObjet = -1 Then
      '  on respecte les couleurs d'origine du fond de plan
    Coul = rechCouleur(couleur, nomPlan, objInsert)
    Coul = QBColor(tCouleur(Abs(Coul)))        ' la couleur peut être négative si plan inactif dans le DXF
  Else
    Coul = coulObjet
  End If
    
  If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
    With objInsert
      Set pa = .TransRotBloc(p)
      R = Rayon * Abs(.echx)
    End With
  Else
    Set pa = p
    R = Rayon
  End If
  
  If pourEchelle Then
    minimax pa.x + R, pa.y + R, pa.x - R, pa.y - R
  Else
    ' dessin
    arcEchelle pa, R, Coul
  End If
  
End Sub

