VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Blocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private colBlocs As New Collection

Public Function Add(ByVal name As String, Optional ByVal position As Integer) As BLOC
' Certains fichiers DXF peuvent avoir un point d'insertion différent de 0,0
' GIRATION translate l'ensemble (Bloc et inserts) en conséquence : le bloc est toujours en 0,0

Dim blocNew As New BLOC
Dim objet As Object
  
  With blocNew
    .name = name
    position = Min(position, colBlocs.Count)
    If position = 0 Then
      colBlocs.Add blocNew, name
    Else
      colBlocs.Add blocNew, name, position
    End If
  End With
  
  Set Add = blocNew
End Function

Public Sub AddBloc(blocNew As BLOC)
    
  colBlocs.Add blocNew, blocNew.name
  
End Sub

Public Function Count() As Long
  Count = colBlocs.Count
End Function

Public Function Item(ByVal Index As Variant) As BLOC
Attribute Item.VB_UserMemId = 0
  Set Item = colBlocs.Item(Index)
End Function

Public Sub Remove(ByVal Index As Variant)
  colBlocs.Remove Index
End Sub

Public Sub Clear()

  Do Until colBlocs.Count <> 0
    colBlocs.Remove 1
  Loop
  
End Sub

' NewEnum doit renvoyer l'interface IUnknown de
' l'énumérateur d'une collection.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colBlocs.[_NewEnum]
End Function

Public Function Exist(ByVal nomBloc As String) As Boolean
Dim wBloc As BLOC

On Error GoTo GestErr
  Set wBloc = colBlocs(nomBloc)
  Exist = True
  Exit Function
  
GestErr:

End Function

  Public Sub Renommer(ByVal nomOrigine As String, ByVal nomNouveau As String)
  Dim BlocOrigine As BLOC, blocNew As BLOC
  Dim p As PT, objPoint As POINT, objLine As LINE, objCercle As CERCLE, objArc As ARC
  Dim pLine As POLYLINE, objInsert As INSERT
  Dim pLineNew As POLYLINE
  
  ' Création du nouveau bloc
  Set blocNew = Add(nomNouveau)
  
  ' Recopie des éléments du bloc origine dans le nouveau bloc
  Set BlocOrigine = colBlocs(nomOrigine)
  
  For Each objPoint In BlocOrigine.blkPoints
    With objPoint
      blocNew.blkPoints.Add .nomPlan, .couleur, .p
    End With
  Next
  For Each objLine In BlocOrigine.blkLines
    With objLine
      blocNew.blkLines.Add .nomPlan, .couleur, .typelign, .Pts(1), .Pts(2)
    End With
  Next
  For Each pLine In BlocOrigine.blkPolylines
    With pLine
      Set pLineNew = blocNew.blkPolylines.Addp(.nomPlan, .couleur, .typelign, .ferme)
      For Each p In .Pts
        pLineNew.Pts.Add p.X, p.Y, p.arrondi
      Next
    End With
  Next
  For Each objCercle In BlocOrigine.blkCercles
    With objCercle
      blocNew.blkCercles.Add .nomPlan, .couleur, .typelign, .p, .Rayon
    End With
  Next
  For Each objArc In BlocOrigine.blkArcs
    With objArc
      blocNew.blkArcs.Add .nomPlan, .couleur, .typelign, .p, .Rayon, .angdep, .angfin
    End With
  Next
  For Each objInsert In BlocOrigine.blkInserts
    With objInsert
      blocNew.blkInserts.Add .name, .nomPlan, .couleur, .typelign, .p, .echx, .echy, .rot, .nbcol, .nbrow, .distcol, .distrow
    End With
  Next
  
  ' Suppression du bloc origine
  colBlocs.Remove nomOrigine

End Sub
