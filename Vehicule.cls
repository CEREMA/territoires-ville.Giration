VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VEHICULE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Public Nom As String
Public Protégé As Boolean
Public nbUtilisation As Integer
Public nbRemorque As Integer
Private mPAFav(2) As Single
Private mEmpat(2) As Single
Private mPAFar(2) As Single
Private mLargeur(2) As Single
Public LEav As Single
Private mLEar(2) As Single
Private mDpat(1) As Single

Public alphaVehMax As Single
Public aVehMax As Single
Public RVehMin As Single

Public Sub Clone(ByVal NewVehic As VEHICULE)
Dim i As Integer

  With NewVehic
    .Nom = Nom
    .nbRemorque = nbRemorque
    For i = 0 To 2
      .PAFav(i) = mPAFav(i)
      .e(i) = mEmpat(i)
      .PAFar(i) = mPAFar(i)
      .l(i) = mLargeur(i)
      .LEar(i) = mLEar(i)
      If i < 2 Then .Dpat(i) = mDpat(i)
    Next
    .LEav = LEav
    .alphaVehMax = alphaVehMax
    .aVehMax = aVehMax
    .RVehMin = RVehMin
  End With
End Sub

Property Get PAFav(ByVal Index As Integer) As Single
  PAFav = mPAFav(Index)
End Property

Property Let PAFav(ByVal Index As Integer, ByVal vNewValue As Single)
  mPAFav(Index) = vNewValue
End Property

Property Get e(ByVal Index As Integer) As Single
  e = mEmpat(Index)
End Property

Property Let e(ByVal Index As Integer, ByVal vNewValue As Single)
  mEmpat(Index) = vNewValue
End Property

Property Get PAFar(ByVal Index As Integer) As Single
  PAFar = mPAFar(Index)
End Property

Property Let PAFar(ByVal Index As Integer, ByVal vNewValue As Single)
  mPAFar(Index) = vNewValue
End Property

Property Get l(ByVal Index As Integer) As Single
  l = mLargeur(Index)
End Property

Property Let l(ByVal Index As Integer, ByVal vNewValue As Single)
  mLargeur(Index) = vNewValue
End Property

Property Get LEar(ByVal Index As Integer) As Single
  LEar = mLEar(Index)
End Property

Property Let LEar(ByVal Index As Integer, ByVal vNewValue As Single)
  mLEar(Index) = vNewValue
End Property

Property Get Dpat(ByVal Index As Integer) As Single
  Dpat = mDpat(Index)
End Property

Property Let Dpat(ByVal Index As Integer, ByVal vNewValue As Single)
  mDpat(Index) = vNewValue
End Property

Public Property Get nomCourt() As String
Dim chaine As String
Dim i As Integer
Dim c As String

On Error GoTo GestErr

  Do While Len(nomCourt) < LGMAXNOMVEHICULE ' 14 est pour limiter le nom de plan à 31 caractères
    i = i + 1
    c = Mid(Nom, i, 1)
    'Ligne suivante : AV Ajout du 16/09/03 - pour éviter la gestion de l'erreur 5
    If c = "" Then Exit Property
    
    Select Case Asc(c)
    ' -(36) $(45) 0-9(48-57) A-Z(65-90) a-z(97-122)
    Case 36, 45, 48 To 57, 65 To 90, 97 To 122
    Case Else ' remplacement du caractère interdit par un tiret
      c = "-"
      If Len(nomCourt) = 0 Or Right(nomCourt, 1) = "-" Then c = "" ' On n'écrit pas 2 tirets de suite
    End Select
    nomCourt = nomCourt & c
  Loop
  'Suppression du séparateur s'il est en fin de mot
  If Right(nomCourt, 1) = "-" Then nomCourt = Left(nomCourt, Len(nomCourt) - 1)

Exit Property

GestErr:
  If Err.Number <> 5 Then ErrGeneral
    'Ligne suivante : AV Ajout du 16/09/03 - pour éviter la gestion de l'erreur 5 (voir ci-dessus)
  ErrGeneral
End Property
