VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colArcs As New Collection

Public Function Add(nomPlan As String, couleur As Long, typelign As String, ByVal centre As PT, Rayon As Double, alpha1 As Double, alpha2 As Double) As ARC
  Dim arcNew As New ARC
  
  With arcNew
  'V18
    .nomPlan = gtPlans(nomPlan).name
  'V20
    .nomPlan = nomPlan
    .couleur = couleur
    .typelign = typelign
    .Rayon = Rayon
    .angdep = alpha1
    .angfin = alpha2
    .p.X = centre.X
    .p.Y = centre.Y
    
    colArcs.Add arcNew
  End With
  
  Set Add = arcNew
End Function

Public Function Count() As Long
  Count = colArcs.Count
End Function

Public Function Item(ByVal Index As Integer) As ARC
Attribute Item.VB_UserMemId = 0
  Set Item = colArcs.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colArcs.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colArcs.[_NewEnum]
End Function


