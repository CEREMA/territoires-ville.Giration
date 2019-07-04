VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public nomPlan As String
Public couleur As Long
Public p As New PT

'****************************************
' Ecriture de l'objet dans un fichier DXF
'****************************************
Public Sub ecrire()

  ecrArticle "POINT", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
    ecr 100, SansCadrage:=True
    ecr "AcDbPoint"
  End If

  ecrPlan nomPlan
  ecrCouleur couleur
  ecrPoint p

End Sub

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
  Dim pa As New PT
  Dim Coul As Long, tlign As Boolean
  
  If coulObjet = -1 Then
      '  on respecte les couleurs d'origine du fond de plan
    Coul = rechCouleur(couleur, nomPlan, objInsert)
    Coul = QBColor(tCouleur(Abs(Coul)))            ' la couleur peut être négative si plan inactif dans le DXF
  Else
    Coul = coulObjet
  End If

  If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
    With objInsert
      Set pa = .TransRotBloc(p)
    End With
  Else
    Set pa = p
  End If

  If pourEchelle Then
    minimax pa.X, pa.Y
  Else
    ' dessin
    ptEchelle pa, Coul
  End If
  
End Sub

