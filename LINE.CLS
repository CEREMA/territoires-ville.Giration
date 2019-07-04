VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public nomPlan As String
Public couleur As Long
Public typelign As String
'Public Points As New Points
Public Pts As New Pts

'****************************************
' Ecriture de l'objet dans un fichier DXF
'****************************************
Public Sub ecrire()

  ecrArticle "LINE", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
    ecr 100, SansCadrage:=True
    ecr "AcDbLine"
  End If

  ecrPlan nomPlan
  ecrCouleur couleur
  ecrTypelign typelign

  With Pts
    ecrPoint .Item(1)
    ecrPoint .Item(2), 1
  End With

End Sub

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
Dim pa As New PT, pb As New PT
Dim Coul As Long
  
  If coulObjet = -1 Then
      '  on respecte les couleurs d'origine du fond de plan
    Coul = rechCouleur(couleur, nomPlan, objInsert)
    Coul = QBColor(tCouleur(Abs(Coul)))                ' la couleur peut être négative si plan inactif dans le DXF
  Else
    Coul = coulObjet
  End If

  If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
    With objInsert
      Set pa = .TransRotBloc(Pts(1))
      Set pb = .TransRotBloc(Pts(2))
    End With
  Else
    Set pa = Pts(1)
    Set pb = Pts(2)
  End If

  If pourEchelle Then
    minimax pa.X, pa.Y, pb.X, pb.Y
  Else
    desEchelle pa, pb, Coul
  End If
  
End Sub

