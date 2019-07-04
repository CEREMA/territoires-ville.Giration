VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colPlans As New Collection

Public Function Add(ByVal unPlan As PLAN) As PLAN
Dim planNew As New PLAN

  With planNew
    .name = unPlan.name
    .couleur = unPlan.couleur
    .typelign = unPlan.typelign
    .gele = unPlan.gele
  End With
  colPlans.Add planNew, unPlan.name
  Set Add = planNew
End Function

Public Function Item(ByVal Index As String) As PLAN
Attribute Item.VB_UserMemId = 0
  Set Item = colPlans(Index)
End Function

Public Function Remove(ByVal Index As String) As PLAN
  colPlans.Remove Index
End Function

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colPlans.[_NewEnum]
End Function

Public Function Count() As Integer
  Count = colPlans.Count
End Function

Public Function Existe(ByVal Nom As String) As Boolean
Dim unPlan As PLAN

On Error GoTo GestErr

  Set unPlan = colPlans(Nom)
  Existe = True
Exit Function

GestErr:
End Function
