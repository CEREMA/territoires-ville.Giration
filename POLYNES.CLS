VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Polylines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colPolylines As New Collection

Public Function Addp(ByVal nomPlan As String, ByVal couleur As Long, _
                    ByVal typelign As String, ByVal ferme As Boolean) As POLYLINE
Dim poligNew As New POLYLINE
  
  With poligNew
  
  'V18
    .nomPlan = gtPlans(nomPlan).name
  'V20
    .nomPlan = nomPlan
    .couleur = couleur
    .typelign = typelign
    .ferme = ferme
    colPolylines.Add poligNew
  End With
  
  Set Addp = poligNew
  
End Function

Public Function AddPline(ByVal pLine As POLYLINE)
  
  colPolylines.Add pLine

  Set AddPline = pLine

End Function

Public Function Count() As Long
  Count = colPolylines.Count
End Function

Public Function Item(ByVal Index As Integer) As POLYLINE
Attribute Item.VB_UserMemId = 0
  Set Item = colPolylines.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colPolylines.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colPolylines.[_NewEnum]
End Function


