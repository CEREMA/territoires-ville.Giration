VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colLines As New Collection

Public Function Add(nomPlan As String, couleur As Long, typelign As String, ByVal p1 As PT, ByVal p2 As PT) As LINE
  Dim ligNew As New LINE
  
  With ligNew
  'V18
    .nomPlan = gtPlans(nomPlan).name
  'V20
    .nomPlan = nomPlan
    .couleur = couleur
    .typelign = typelign
    .Pts.Add p1.X, p1.Y
    .Pts.Add p2.X, p2.Y
    
    colLines.Add ligNew
  End With
  
  Set Add = ligNew

End Function

Public Function Count() As Long
  Count = colLines.Count
End Function

Public Function Item(ByVal Index As Integer) As LINE
Attribute Item.VB_UserMemId = 0
  Set Item = colLines.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colLines.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colLines.[_NewEnum]
End Function


