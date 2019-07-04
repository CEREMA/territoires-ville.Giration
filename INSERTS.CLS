VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colInserts As New Collection

Public Function Add(name As String, nomPlan As String, couleur As Long, typelign As String, ByVal p1 As PT, echx As Double, echy As Double, rot As Double, nbcol As Integer, nbrow As Integer, distcol As Double, distrow As Double) As INSERT
  Dim insertNew As New INSERT
                                                      
  
  With insertNew
    .name = name
  'V18
    .nomPlan = gtPlans(nomPlan).name
  'V20
    .nomPlan = nomPlan
    .couleur = couleur
    .typelign = typelign
    .p.X = p1.X
    .p.Y = p1.Y
    .echx = echx
    .echy = echy
    .rot = rot
    .nbcol = nbcol
    .nbrow = nbrow
    .distcol = distcol
    .distrow = distrow
    colInserts.Add insertNew
  End With
  
  Set Add = insertNew
End Function

Public Function Count() As Long
  Count = colInserts.Count
End Function

Public Function Item(ByVal Index As Integer) As INSERT
Attribute Item.VB_UserMemId = 0
  Set Item = colInserts.Item(Index)
End Function

Public Sub Remove(ByVal Index As Integer)
  colInserts.Remove Index
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colInserts.[_NewEnum]
End Function


