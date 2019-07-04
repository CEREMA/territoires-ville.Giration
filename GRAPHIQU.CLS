VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRAPHIQUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public gphPoints As New Points
Public gphLines As New Lines
Public gphPolylines As New Polylines
Public gphSplines As New Splines
Public gphCercles As New Cercles
Public gphArcs As New Arcs
Public gphInserts As New Inserts
Public gphBlocs As New Blocs

'*************************************************************
' Retourne le nombre d'éléments composants le bloc
'*************************************************************
Public Function nbElements() As Integer
  
  nbElements = gphPoints.Count + gphLines.Count + gphPolylines.Count + gphSplines.Count
  nbElements = nbElements + gphCercles.Count + gphArcs.Count + gphInserts.Count
  
End Function

