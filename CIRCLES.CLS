VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cercles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colCercles As New Collection

Public Function Add(ByVal nomPlan As String, ByVal couleur As Long, ByVal typelign As String, ByVal centre As PT, ByVal Rayon As Double) As CERCLE
  Dim cercleNew As New CERCLE
  
  With cercleNew
  'V18
    .nomPlan = gtPlans(nomPlan).name
  'V20
    .nomPlan = nomPlan
    .couleur = couleur
    .typelign = typelign
    .Rayon = Rayon
    .p.X = centre.X
    .p.Y = centre.Y
    
    colCercles.Add cercleNew
  End With
  
  Set Add = cercleNew
  
End Function

Public Function Count() As Long
  Count = colCercles.Count
End Function

Public Function Item(ByVal Index As Integer) As CERCLE
Attribute Item.VB_UserMemId = 0
  Set Item = colCercles.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colCercles.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colCercles.[_NewEnum]
End Function


