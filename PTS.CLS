VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colPts As New Collection

Public Function Add(ByVal X1 As Double, ByVal Y1 As Double, Optional ByVal arrondi As Double, Optional ByVal Drapeau As Integer) As PT
Dim ptNew As New PT
  
  With ptNew
    .X = X1
    .Y = Y1
    .arrondi = arrondi
    .Drapeau = Drapeau
  End With
  
  colPts.Add ptNew
  Set Add = ptNew
  
End Function

Public Function AddPt(p As PT) As PT
Dim ptNew As New PT
  
  With ptNew
    .X = p.X
    .Y = p.Y
    .arrondi = p.arrondi
    .Drapeau = p.Drapeau
  End With
  
  colPts.Add ptNew
  Set AddPt = ptNew
  
End Function

Public Function Count() As Long
  Count = colPts.Count
End Function

Public Function Item(ByVal Index As Integer) As PT
Attribute Item.VB_UserMemId = 0
  Set Item = colPts.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colPts.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colPts.[_NewEnum]
End Function

