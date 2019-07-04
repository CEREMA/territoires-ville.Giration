VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Points"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colPoints As New Collection

Public Function Add(nomPlan As String, couleur As Long, ByVal p1 As PT) As POINT
  Dim ptNew As New POINT
  
  With ptNew
  'V18
    .nomPlan = gtPlans(nomPlan).name
  'V20
    .nomPlan = nomPlan
    .couleur = .couleur
    .p.X = p1.X
    .p.Y = p1.Y
    colPoints.Add ptNew
  End With
  
  Set Add = ptNew

End Function

Public Function Count() As Long
  Count = colPoints.Count
End Function

Public Function Item(ByVal Index As Integer) As POINT
Attribute Item.VB_UserMemId = 0
  Set Item = colPoints.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colPoints.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colPoints.[_NewEnum]
End Function


