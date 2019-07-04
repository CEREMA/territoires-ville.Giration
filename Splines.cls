VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Splines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colSplines As New Collection

Public Function Add(ByVal maSpLine As SPLINE)
  
  colSplines.Add maSpLine

  Set Add = maSpLine

End Function

Public Function Count() As Long
  Count = colSplines.Count
End Function

Public Function Item(ByVal Index As Integer) As SPLINE
Attribute Item.VB_UserMemId = 0
  Set Item = colSplines.Item(Index)
End Function

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colSplines.[_NewEnum]
End Function
