VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INSERT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public name As String
Public nomPlan As String
Public couleur As Long
Public typelign As String
Public p As New PT
'Public Pts As New Pts
Public echx As Double, echy As Double, rot As Double, nbcol As Integer, nbrow As Integer, distcol As Double, distrow As Double

'****************************************
' Ecriture de l'objet dans un fichier DXF
'****************************************
Public Sub ecrire()

  ecrArticle "INSERT", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
    ecr 100, SansCadrage:=True
    ecr "AcDbBlockReference"
  End If
  
  ecrPlan nomPlan
  ecrCouleur couleur
  ecrTypelign typelign
  ecrNom name
  ecrPoint p
  If echx <> 1 Then   ' GIRATION n'accepte que les facteurs d'échelle égaux en X et Y
    ecr 41, True
    ecr echx
    ecr 42, True
    ecr echx
  End If
  If rot <> 0 Then
    ecr 50, True
    If OptGen.unite = GRADE Then
      ecr rot * 0.9
    Else
      ecr rot
    End If
  End If
  If nbcol <> 1 Then
    ecr 70
    ecr nbcol
    ecr 44
    ecr distcol
  End If
  If nbrow <> 1 Then
    ecr 70
    ecr nbrow
    ecr 44
    ecr distrow
  End If
  
  
End Sub

Public Sub dessiner(ByVal coulObjet As Long, Optional ByVal objInsert As INSERT)
Dim nomBloc As BLOC
Dim objetPoint As POINT
Dim objetLine As LINE
Dim objetArc As ARC
Dim objetCercle As CERCLE
Dim objetPolyline As POLYLINE
Dim objetSpline As SPLINE
Dim objetInsert As INSERT
Dim Cpt As Integer
Dim xa As Double, ya As Double
Dim pa As PT

If Not objInsert Is Nothing Then              ' entité insérée dans 1 bloc
  Set objetInsert = New INSERT                ' Blocs imbriqués: cf Manuel utilisateur AutoCAD 12 ch.10 p 394...
  With objInsert
    Set pa = .TransRotBloc(p)
  End With
  With objetInsert
    .name = name
    If nomPlan = "0" Then    'l'insert étant dans le plan "0", on crée l'objet dans le plan du 'père'
      .nomPlan = gtPlans(objInsert.nomPlan).name
    Else
      .nomPlan = nomPlan
    End If
    If couleur = 0 Then 'l'insert étant de la couleur DUBLOC on lui donne la couleur du 'père'
      .couleur = rechCouleur(couleur, nomPlan, objInsert)
    Else
      .couleur = couleur
    End If
    If typelign = "BYBLOCK" Then 'l'insert étant du type de ligne DUBLOC on lui donne le type de ligne du 'père'
      .typelign = gtPlans(objInsert.nomPlan).typelign
    Else
      .typelign = typelign
    End If
    Set .p = pa
    .echx = echx * objInsert.echx
    'Ajout AV( 14/11/06 : v 3.3.1) réserve l'avenir
    .echy = echy * objInsert.echy
    .rot = rot + objInsert.rot
    .dessiner coulObjet
  End With
  Set objetInsert = Nothing
    
Else
  
  If flagVéhicule Then
    Set nomBloc = fCourante.Carross.gphBlocs(name)
    With nomBloc.blkPolylines
      For Cpt = 1 To .Count
        .Item(Cpt).dessiner coulObjet, Me
      Next
    End With
  Else
    
    Set nomBloc = fCourante.GraphFDP.gphBlocs(name)
    
    With nomBloc
      For Each objetPoint In .blkPoints
        objetPoint.dessiner coulObjet, Me
      Next
      For Each objetLine In .blkLines
        objetLine.dessiner coulObjet, Me
      Next
      For Each objetPolyline In .blkPolylines
        objetPolyline.dessiner coulObjet, Me
      Next
      For Each objetSpline In .blkSplines
        objetSpline.dessiner coulObjet, Me
      Next
      For Each objetArc In .blkArcs
        objetArc.dessiner coulObjet, Me
      Next
      For Each objetCercle In .blkCercles
        objetCercle.dessiner coulObjet, Me
      Next
      For Each objetInsert In .blkInserts
        objetInsert.dessiner coulObjet, Me
      Next
    End With    ' nombloc
    
  End If                  ' flagVéhicule
  
End If                    ' bloc inséré dans 1 bloc

End Sub

'**************************************************************************************
' Translation-Rotation du point d'insertion d'un bloc avec ou sans facteur d'échelle
'**************************************************************************************
Public Function TransRotBloc(ByVal unPoint As PT) As PT
  
  Set TransRotBloc = TransRot(unPoint, p, rot, echx)

End Function

Private Sub Class_Initialize()
  echx = 1
  echy = 1
  nbcol = 1
  nbrow = 1
End Sub
