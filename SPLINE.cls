VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPLINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public nomPlan As String
Public couleur As Long
Public typelign As String
Public ferme As Boolean

Public Ordre As Integer
Public Rational As Boolean

Public tolNoeuds As Single
Public tolPControle As Single

Public PtsControle As New Pts
Public PtsDessinés As New Pts

Private mNoeuds() As Single

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
Dim i As Integer, j As Integer, nbPoints As Integer, nbSegments As Integer, iSuivant As Integer
Dim pa As PT, pb As PT
Dim Coul As Long
Dim mPts As Pts
  
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

  For i = 1 To PtsDessinés.Count
    
    If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
      With objInsert
        Set pb = .TransRotBloc(PtsDessinés(i))
      End With
    Else
      Set pb = PtsDessinés(i)
    End If
  
    If pourEchelle Then
      minimax pb.X, pb.Y
    ElseIf i > 1 Then
      desEchelle pa, pb, Coul
    End If
    Set pa = pb
  Next

  
End Sub

Public Sub Etablir(ByVal nbPoints As Integer)
Dim fin As Single, intervalle As Single
Dim i As Integer, j As Integer
Dim temp As Single
Dim nbasis() As Double
Dim p As PT
Dim pb As New PT
Dim nPlusOrdre As Integer
Dim t As Single
Dim npts As Integer

  'PtsDessinés
  
  npts = PtsControle.Count
  
  intervalle = Noeuds(UBound(mNoeuds)) / (nbPoints - 1)

  ReDim nbasis(1 To npts)
  'Il faut n fonctions d'ordre k(=Ordre) que de points de controle
  'Il faut n+k fonctions de base d'ordre 1 puisqu'on perd une fonction à chaque fois qu'on élève le degré et donc l'ordre de la courbe
  nPlusOrdre = npts + Ordre
  
  For i = 1 To nbPoints

    If Noeuds(nPlusOrdre - 1) - t < 0.000005 Then
      t = Noeuds(nPlusOrdre - 1)
    End If

    Basis t, nbasis    '      /* generate the basis function for this value of t */
    
    pb.X = 0#
    pb.Y = 0#
    j = 0
    For Each p In PtsControle    ' /* Do local matrix multiplication */
      j = j + 1
      temp = nbasis(j) * p.X
      pb.X = pb.X + temp
      temp = nbasis(j) * p.Y
      pb.Y = pb.Y + temp
    Next
    
    PtsDessinés.Add pb.X, pb.Y
      
    t = t + intervalle
  Next

End Sub

Private Sub Basis(ByVal t As Single, ByRef n)

'/*  Subroutine to generate B-spline basis functions for open knot vectors
'    t        = parameter value
'    n[]      = array containing the basis functions
'               n[1] contains the basis function associated with B1 etc.

'    nPlusOrdre   = constant -- npts + ordre -- maximum number of knot values
'    temp[]   = temporary array

'    d        = first term of the basis function recursion relation
'    e        = second term of the basis function recursion relation
'*/

Dim nPlusOrdre As Integer
Dim i As Integer, k As Integer
Dim d As Double, e As Double
Dim temp() As Double
Dim Somme As Single

  nPlusOrdre = PtsControle.Count + Ordre

  ReDim temp(1 To nPlusOrdre)

'/* calculate the first order basis functions n[i][1]  */

  For i = 1 To nPlusOrdre - 1
    If t >= Noeuds(i - 1) And t < Noeuds(i) Then
      temp(i) = 1
    Else
      temp(i) = 0
    End If
 Next

'/* calculate the higher order basis functions */

  For k = 2 To Ordre
    For i = 1 To nPlusOrdre - k
      If temp(i) = 0# Then   '    /* if the lower order basis function is zero skip the calculation */
        d = 0
      Else
        d = ((t - Noeuds(i - 1)) * temp(i)) / (Noeuds(i + k - 2) - Noeuds(i - 1))
      End If

      If temp(i + 1) = 0# Then  '     /* if the lower order basis function is zero skip the calculation */
        e = 0
      Else
        e = ((Noeuds(i + k - 1) - t) * temp(i + 1)) / (Noeuds(i + k - 1) - Noeuds(i))
      End If

      temp(i) = d + e
    Next
  Next

  If (t = Noeuds(nPlusOrdre - 1)) Then '   /*    pick up last point  */
    temp(PtsControle.Count) = 1
  End If

'/* put in n array */

  If Rational Then
    Somme = 0#
    For i = 1 To PtsControle.Count
      Somme = Somme + temp(i) * PtsControle(i).arrondi
    Next
    For i = 1 To PtsControle.Count
      If Somme <> 0# Then
        n(i) = temp(i) * PtsControle(i).arrondi / Somme
      Else
        n(i) = 0#
      End If
    Next
  
  Else
    For i = 1 To PtsControle.Count
      n(i) = temp(i)
    Next
  End If

End Sub

Public Property Get Noeuds(ByVal Index As Integer) As Single
  Noeuds = mNoeuds(Index)
End Property

Public Property Let Noeuds(ByVal Index As Integer, ByVal vNewValue As Single)
  If Index > UBound(mNoeuds) Then ReDim Preserve mNoeuds(Index)
  mNoeuds(Index) = vNewValue
End Property

Private Sub Class_Initialize()
  ReDim mNoeuds(0)
End Sub

Public Sub ecrire()
Dim intValeur As Integer
Dim p As PT
Dim i As Integer

  ecrArticle "SPLINE", AvecID:=True
  ecr 100, SansCadrage:=True
  ecr "AcDbEntity"
  ecr 100, SansCadrage:=True
  ecr "AcDbSp line"
    
  ecrPlan nomPlan
  ecrCouleur couleur
  ecrTypelign typelign

  ecr 70, SansCadrage:=True
  intValeur = 8
  If ferme Then intValeur = intValeur + 1
  If Rational Then intValeur = intValeur + 4
  ecr intValeur
  
  ecr 71, SansCadrage:=True
  ecr Ordre - 1
  
  ecr 72, SansCadrage:=True
  ecr UBound(mNoeuds) + 1
  
  ecr 73, SansCadrage:=True
  ecr PtsControle.Count
  
  ecr 42, SansCadrage:=True
  ecr tolNoeuds
  
  ecr 43, SansCadrage:=True
  ecr tolPControle
  
  For i = 0 To UBound(mNoeuds)
    ecr 40, SansCadrage:=True
    ecr mNoeuds(i)
  Next
  
  For Each p In PtsControle
    ecrPoint p
    If Rational Then
      ecr 41, SansCadrage:=True
      ecr p.arrondi
    End If
  Next
  
End Sub
