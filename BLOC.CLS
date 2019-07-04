VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BLOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  
Public name As String
'Public plan As String        non significatif dans le DXF (plan 0)
'Public p As New PT           'en principe toujours 0,0,0 dans les DXF
'Public couleur As Long    non   informé       dans le DXF
'Public typelign as String      "         "              "           "
Public blkPoints As New Points
Public blkLines As New Lines
Public blkPolylines As New Polylines
Public blkSplines As New Splines
Public blkCercles As New Cercles
Public blkArcs As New Arcs
Public blkInserts As New Inserts

Public IdBlockTableRecord As String

'*************************************************************
' Retourne le nombre d'éléments composants le bloc
'*************************************************************
Public Function nbElements() As Integer
  
  nbElements = blkPoints.Count + blkLines.Count + blkPolylines.Count + blkSplines.Count
  nbElements = nbElements + blkCercles.Count + blkArcs.Count + blkInserts.Count
  
End Function
