Attribute VB_Name = "EcrDXF"
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module standard : EcrDXF   -   ECRDXF.BAS
'   Ecriture d'un fichier DXF
'       soit un fond de plan
'       soit export d'une épure (dans ce cas, flagExport = True)

'**************************************************************************************

Option Explicit
Private EcrireFDP As Boolean
Private basGauche As New PT, hautDroit As New PT
Private numFich As Integer
Private nomCourtVéhicule As String
Private IDAcad As Long
Private MaxHandle As String
Private strIDTable As String

Private mBlocs As Blocs
Private nbPlans As Integer
Public VersionDXF As Integer

Private Sub ecrBloc(ByVal nBloc As BLOC, Optional ByVal IdBlock_Record As String)
Dim p As New PT

' Tous les blocs dans le DXF sont référencés en 0,0
'(même si certains DXF en ont un, GIRATION effectue en lecture les translations nécessaires : voir fonction lire_entite dans LECDXF.BAS)
  p.x = 0
  p.y = 0
    
  ecrEnteteBloc nBloc.IdBlockTableRecord, EspacePapier:=nBloc.name = "*PAPER_SPACE"
  ecrNom nBloc.name
  ecr 70, True
  If nBloc.name = "*MODEL_SPACE" Or nBloc.name = "*PAPER_SPACE" Then
    ecr 0
  Else
    ecr 64            ' bloc référencé dans le dessin
  End If
  ecrPoint p
  ecr 3, True
  ecr FormatNomV14(nBloc.name)
  ecr_listentites nBloc
  ecrFinBloc EspacePapier:=nBloc.name = "*PAPER_SPACE"

End Sub

Public Sub EcrDXF(ByVal NomFich As String)
nomCourtVéhicule = fCourante.Vehic.nomCourt

#If V2000 Then
  VersionDXF = 15       ' 2000 ou 2002
#ElseIf V14 Then
  VersionDXF = 14
#Else
  VersionDXF = 11       ' v11 ou 12
#End If

  VersionDXF = 14
  
  EcrireFDP = fCourante.FDP <> "" And (Not flagExport Or OptionAffichage(NumFille).FDP.ADessiner)
  ' Fonction Enregistrer le Fond de plan ou
  ' Fonction Export avec Option d'affichage du FDP activée
          
  ActivMenu Activation:=False
  Screen.MousePointer = vbHourglass    ' Sablier
  
  With fCourante
    basGauche.x = .pmin.x
    basGauche.y = .pmin.y
    hautDroit.x = .pmax.x
    hautDroit.y = .pmax.y
          
    If EcrireFDP Then
      Wpmin.x = 1000000000#
      Wpmin.y = 1000000000#
      Wpmax.x = -1000000000#
      Wpmax.y = -1000000000#
      calminmax
      .pminFDP.x = Wpmin.x
      .pminFDP.y = Wpmin.y
      .pmaxFDP.x = Wpmax.x
      .pmaxFDP.y = Wpmax.y
    End If
    flagEcranZoom = True    ' pour que ToutVoir et CalcEchelle se contentent de calculer les limites
    If flagExport Then
      flagExport = False
      ToutVoir
      flagExport = True
      ActivMenu Activation:=False
    Else
      CalcEchelle
    End If
    flagEcranZoom = False
    
  End With    ' Fin du traitement final
            
  CalculerNbPlans
  InstancierBlocs
    
  numFich = FreeFile
  ' Ecrasement de l'ancien fichier(si existe)
  Open NomFich For Output As numFich
  Close numFich
  
  Open NomFich For Binary Lock Read Write As numFich
  
  ecrArticle "SECTION"
  ecrNom "HEADER"
  ecrireHeader
  ecrArticle "ENDSEC"
  
  ecrArticle "SECTION"
  ecrNom "TABLES"
  ecrireTables
  ecrArticle "ENDSEC"
  
  ecrArticle "SECTION"
  ecrNom "BLOCKS"
  ecrireBlocs
  ecrArticle "ENDSEC"
  
  ecrArticle "SECTION"
  ecrNom "ENTITIES"
  If EcrireFDP Then
    ecr_listentites fCourante.GraphFDP
  End If
  If flagExport Then
    ' Fonction Export
    ecr_listentites fCourante.GraphExport
  End If
  ecrArticle "ENDSEC"

  If VersionDXF > 11 Then
    ecrArticle "SECTION"
    ecrNom "OBJECTS"
    ecrireDictionnaires
    ecrArticle "ENDSEC"
    IDAcad = 0
  End If
  
  ecrArticle "EOF"
  
  Close numFich
  
  mBlocs.Clear
  Set mBlocs = Nothing

  ActivMenu Activation:=True
  Screen.MousePointer = vbDefault

End Sub

Private Sub InstancierBlocs()
Dim nBloc As BLOC
  
  Set mBlocs = New Blocs
  
  If EcrireFDP Then

    With fCourante.GraphFDP
        ' Blocs du Fond de plan
      For Each nBloc In .gphBlocs
        mBlocs.AddBloc nBloc
      Next
    
    End With  ' fCourante.GraphFDP
  End If

  If flagExport Then
    With fCourante.GraphExport
      ' Blocs des carrosseries
      For Each nBloc In .gphBlocs
        mBlocs.AddBloc nBloc
      Next
    End With
  End If

End Sub

Private Sub ecr_listentites(ByVal objet As Object)
Dim Cpt%
''   objet = fCourante.GraphFDP    Fond de plan
''ou         fCourante.GraphExport si export
''ou un bloc BLOC
Dim wBloc As BLOC
Dim wGraph As GRAPHIQUE

Dim objPoint As POINT
Dim objLine As LINE
Dim objPline As POLYLINE
Dim objArc As ARC
Dim objCercle As CERCLE
Dim objSpline As SPLINE
Dim objInsert As INSERT

  If TypeOf objet Is BLOC Then
    Set wBloc = objet
    With wBloc
      For Each objPoint In .blkPoints
        objPoint.ecrire
      Next
      
      For Each objLine In .blkLines
        objLine.ecrire
      Next
      
      For Each objPline In .blkPolylines
        objPline.ecrire
      Next
      
      For Each objSpline In .blkSplines
        objSpline.ecrire
      Next
      
      For Each objCercle In .blkCercles
        objCercle.ecrire
      Next
      
      For Each objArc In .blkArcs
        objArc.ecrire
      Next
      
      For Each objInsert In .blkInserts
        objInsert.ecrire
      Next
    
    End With  ' objet
    
  Else
      
    Set wGraph = objet
    With wGraph
      For Each objPoint In .gphPoints
        objPoint.ecrire
      Next
      
      For Each objLine In .gphLines
        objLine.ecrire
      Next
      
      For Each objPline In .gphPolylines
        objPline.ecrire
      Next
      
      For Each objSpline In .gphSplines
        objSpline.ecrire
      Next
      
      For Each objCercle In .gphCercles
        objCercle.ecrire
      Next
      
      For Each objArc In .gphArcs
        objArc.ecrire
      Next
      
      For Each objInsert In .gphInserts
        objInsert.ecrire
      Next
    
    End With  ' objet
  End If
  
End Sub

Private Sub ecrireBlocs()
Dim nBloc As BLOC

  For Each nBloc In mBlocs
    ecrBloc nBloc
  Next

End Sub

Private Sub ecrFinBloc(ByVal EspacePapier As Boolean)
  
  ecrArticle "ENDBLK", AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
    If EspacePapier Then
      ecr 67, SansCadrage:=True
      ecr 1
    End If
    ecr 100, SansCadrage:=True
    ecr "AcDbBlockEnd"
  End If

End Sub

Private Sub CalculerNbPlans()

  If EcrireFDP Then
    nbPlans = fCourante.colPlans.Count - (4 + fCourante.nbRemorque * 2)
  Else
    nbPlans = 1   ' Plan "0" : nécessaire en v14
  End If
  
  If flagExport Then nbPlans = nbPlans + ExportPlans.Count
  
End Sub

Private Sub ecrTableLayer()

Dim objPlan As New PLAN

  ecrFamilleTable "LAYER", nbMaxTables:=nbPlans

  If EcrireFDP Then
    For Each objPlan In fCourante.colPlans
      If flagExport And ExportPlans.Existe(objPlan.name) Then
      Else
        ecr_plan objPlan
      End If
    Next
  Else
 ' Plan "0" : nécessaire en v14
    ecr_plan objPlan:=Nothing
  End If

  If flagExport Then
    For Each objPlan In ExportPlans
      ecr_plan objPlan
    Next
  End If
  
  ecrArticle "ENDTAB"

End Sub

Private Sub ecr_plan(ByVal objPlan As PLAN)
Dim chaine As String

  If objPlan Is Nothing Then
    Set objPlan = New PLAN
    With objPlan
      .name = "0"
      .couleur = 7
      .typelign = "CONTINUOUS"
    End With
  End If

  With objPlan
    If InStr(1, .name, "$$$GIR_") = 0 Then

      ecrEnteteTable "LAYER", "AcDbLayerTableRecord"
      ecrNom .name
      ecr 70, True
      ecr 0
          ' Plan réservé GIRATION
      If InStr(.name, "GIR_") = 1 And InStr(5, .name, nomCourtVéhicule, vbTextCompare) = 5 Then
        ecrCouleur acadCol(.couleur)
      Else
        ecrCouleur .couleur
      End If
      ecrTypelign .typelign
      If VersionDXF = 15 Then
        ecr 390, SansCadrage:=True
        ecr MaxHandle
      End If
    End If
  End With
  
End Sub

Private Sub ecrTableViewport()

'CCC  IDAcad = IDAcad + 1

  ecrFamilleTable "VPORT", nbMaxTables:=1  ' 2 ??

'CCC  IDAcad = IDAcad + 1
  ecrEnteteTable "VPORT", "AcDbViewportTableRecord"
  
  ecrNom "*ACTIVE"
  ecr 70, True
  ecr 0
  ecr 10, True
  ecr 0#
  ecr 20, True
  ecr 0#
  ecr 11, True
  ecr 1#
  ecr 21, True
  ecr 1#
  ecr 12, True
  ecr 276.263599
  ecr 22, True
  ecr 166.302521
  ecr 13, True
  ecr 0#
  ecr 23, True
  ecr 0#
  ecr 14, True
  ecr 1#
  ecr 24, True
  ecr 1#
  ecr 15, True
  ecr 10#
  ecr 25, True
  ecr 10#
  ecr 16, True
  ecr 0#
  ecr 26, True
  ecr 0#
  ecr 36, True
  ecr 1#
  ecr 17, True
  ecr 0#
  ecr 27, True
  ecr 0#
  ecr 37, True
  ecr 0#
  ecr 40, True
  ecr 332.605043
  ecr 41, True
  ecr 1.661211
  ecr 42, True
  ecr 50#
  ecr 43, True
  ecr 0#
  ecr 44, True
  ecr 0#
  ecr 50, True
  ecr 0#
  ecr 51, True
  ecr 0#
  ecr 71, True
  ecr 0
  ecr 72, True
  ecr 100
  ecr 73, True
  ecr 1
  ecr 74, True
  ecr 0       ' Aucune icone affichée (ICONESCU)
  ecr 75, True
  ecr 0
  ecr 76, True
  ecr 0
  ecr 77, True
  ecr 0
  ecr 78, True
  ecr 0

    ' insérer ici d'autres VPORT
    
  ecr 0, True
  ecr "ENDTAB"

End Sub

Private Sub ecrEnteteTable(ByVal nomTable As String, ByVal strTableRecord As String)

  ecrArticle nomTable, AvecID:=VersionDXF > 11
  
  If VersionDXF > 11 Then
    ecr 330, SansCadrage:=True
    ecr strIDTable
    ecr 100, SansCadrage:=True
    ecr "AcDbSymbolTableRecord"
    ecr 100, SansCadrage:=True
    ecr strTableRecord
  End If

End Sub

Private Sub ecrEnteteBloc(ByVal IdBlock_Record As String, ByVal EspacePapier As Boolean)
  
  ecrArticle "BLOCK", AvecID:=VersionDXF > 11

  If VersionDXF > 11 Then
    ecr 330, SansCadrage:=True
    ecr IdBlock_Record
    ecr 100, SansCadrage:=True
    ecr "AcDbEntity"
  
    If EspacePapier Then
      ecr 67, SansCadrage:=True
      ecr 1
    End If
    
    ecr 100, SansCadrage:=True
    ecr "AcDbBlockBegin"
  End If

End Sub

Private Sub ecrFamilleTable(ByVal nomTable As String, ByVal nbMaxTables As Integer)

  ecrArticle "TABLE"
  ecrNom nomTable

  If VersionDXF > 11 Then
    strIDTable = ecrNewID(CodeID:=5)       ' Incrémente IDAcad et écrit l'ID (code 5)
    ecr 100, SansCadrage:=True
    ecr "AcDbSymbolTable"
  End If
  
  ecr 70, SansCadrage:=True
  ecr nbMaxTables       ' nombre maxi de Tables
  
End Sub

Private Sub ecrTableLtype()

  ecrFamilleTable "LTYPE", nbMaxTables:=1

  If VersionDXF > 11 Then
    ecrEnteteTable "LTYPE", "AcDbLinetypeTableRecord"
  
    ecrNom "BYBLOCK"
    ecr 70, SansCadrage:=True
    ecr 0
    ecr 3, True
    ecr ""
    ecr 72, True
    ecr 65
    ecr 73, True
    ecr 0
    ecr 40, True
    ecr 0#
  
    ecrEnteteTable "LTYPE", "AcDbLinetypeTableRecord"
    ecrNom "BYLAYER"
    ecr 70, SansCadrage:=True
    ecr 0
    ecr 3, True
    ecr ""
    ecr 72, True
    ecr 65
    ecr 73, True
    ecr 0
    ecr 40, True
    ecr 0#
  End If
  
  ecrEnteteTable "LTYPE", "AcDbLinetypeTableRecord"
  ecrNom "CONTINUOUS"
  ecr 70, SansCadrage:=True
  ecr 0
  ecr 3, True
  ecr "Solid line"
  ecr 72, True
  ecr 65
  ecr 73, True
  ecr 0
  ecr 40, True
  ecr 0#

' debug : si on rajoute le type POINTILLE, il faudra (?) incrémenter le nombre maxi de type de lignes + haut
'ecrArticle "LTYPE"
'ecrNom "POINTILLE"
'ecr 70, True
'ecr 64
'ecr 3, True
'ecr ". . . . . . . . . . . . . . . . . . . . . . . ."
'ecr 72, True
'ecr 65
'ecr 73, True
'ecr 2
'ecr 40, True
'ecr 0.25
'ecr 49, True
'ecr 0#
'ecr 49, True
'ecr -0.25

ecrArticle "ENDTAB"

End Sub

Private Sub ecrTableStyle()

  ecrFamilleTable "STYLE", nbMaxTables:=1
    
  ecrEnteteTable "STYLE", "AcDbTextStyleTableRecord"

  ecrNom "STANDARD"
  ecr 70, True
  ecr 0
  ecr 40, True
  ecr 0#
  ecr 41, True
  ecr 1#
  ecr 50, True
  ecr 0#
  ecr 71, True
  ecr 0
  ecr 42, True
  ecr 3.5         ' v14 : 2.5 ?
  ecr 3, True
  ecr "txt"
  ecr 4, True
  ecr ""
  
  ecrArticle "ENDTAB"

End Sub

Private Sub ecrTablesInutilisées()

  ' Table des vues
  ecrFamilleTable "VIEW", nbMaxTables:=0
  ecrArticle "ENDTAB"
  
  ' Tables des SCU
  ecrFamilleTable "UCS", nbMaxTables:=0
  ecrArticle "ENDTAB"
  
  'Table des appli  ' une seule application : ACAD
  ecrFamilleTable "APPID", nbMaxTables:=1
  ecrEnteteTable "APPID", "AcDbRegAppTableRecord"
  ecrNom "ACAD"
  ecr 70, True
  ecr 0
  ecrArticle "ENDTAB"
  
  ' Tables des Styles de cotation
  ecrFamilleTable "DIMSTYLE", nbMaxTables:=0
  If VersionDXF = 15 Then
    ecr 100, SansCadrage:=True
    ecr "AcDbDimStyleTable"
    ecr 71, SansCadrage:=True
    ecr 0
  End If
  ecrArticle "ENDTAB"

End Sub

Private Sub ecrTableBloc()
Dim nBloc As BLOC

On Error GoTo GestErr

  ' Tables des BLOCK_RECORD (Espace Objet et Papier)
  ecrFamilleTable "BLOCK_RECORD", nbMaxTables:=2 + mBlocs.Count

  Set nBloc = mBlocs.Add("*MODEL_SPACE", position:=1)

  Set nBloc = mBlocs.Add("*PAPER_SPACE", position:=2)
  
  For Each nBloc In mBlocs
    ecrEnteteTable "BLOCK_RECORD", "AcDbBlockTableRecord"
    ecrNom nBloc.name
    nBloc.IdBlockTableRecord = Hex(IDAcad)
  Next

  ecrArticle "ENDTAB"

Exit Sub
GestErr:
  Resume
End Sub

Private Sub ecrireHeader()
Dim p  As New PT

ecrVariable "$ACADVER"
ecr 1, True
Select Case VersionDXF
Case 11
  ecr "AC1009"      'Acad R11 / R12
Case 14
  ecr "AC1014"      'Acad R14
Case 15
  ecr "AC1015"      'Acad 2000/2002
End Select

ecrVariable "$INSBASE"
ecrPoint p

ecrVariable "$EXTMIN"       ' extensions
ecrPoint Wpmin

ecrVariable "$EXTMAX"
ecrPoint Wpmax

ecrVariable "$LIMMIN"   ' limites
ecrPoint p

ecrVariable "$LIMMAX"
p.x = 410
p.y = 287
ecrPoint p

ecrVariable "$ORTHOMODE"
ecr 70, True
ecr 0

ecrVariable "$REGENMODE"
ecr 70, True
ecr 1

ecrVariable "$FILLMODE"
ecr 70, True
ecr 1

ecrVariable "$QTEXTMODE"
ecr 70, True
ecr 0

ecrVariable "$MIRRTEXT"
ecr 70, True
ecr 1

ecrVariable "$LTSCALE"
ecr 40, True
ecr 25#

ecrVariable "$ATTMODE"
ecr 70, True
ecr 1

ecrVariable "$TEXTSIZE"
ecr 40, True
ecr 3.5

ecrVariable "$TRACEWID"
ecr 40, True
ecr 1#

ecrVariable "$TEXTSTYLE"
ecr 7, True
ecr "STANDARD"

ecrVariable "$CLAYER"
ecr 8, True
ecr "0"

ecrVariable "$CELTYPE"
ecr 6, True
ecr "BYLAYER"

ecrVariable "$CECOLOR"
ecr 62, True
ecr 256

ecrVariable "$LUNITS"
ecr 70, True
ecr 2

ecrVariable "$LUPREC"
ecr 70, True
ecr 2

ecrVariable "$SKETCHINC"
ecr 40, True
ecr 1#

ecrVariable "$FILLETRAD"
ecr 40, True
ecr 0#

ecrVariable "$AUNITS"     ' unité d'angle   0 (ou 2 pour les grades)
ecr 70, True
If OptGen.unite = GRADE Then
  ecr 2
Else
  ecr 0
End If

ecrVariable "$AUPREC"
ecr 70, True
ecr 0

ecrVariable "$ELEVATION"
ecr 40, True
ecr 0#

ecrVariable "$PELEVATION"
ecr 40, True
ecr 0#

ecrVariable "$THICKNESS"
ecr 40, True
ecr 0#

ecrVariable "$LIMCHECK"
ecr 70, True
ecr 0

ecrVariable "$SKPOLY"
ecr 70, True
ecr 0

ecrVariable "$TDCREATE"           'Dates : 2450450  corresspond au 1/01/1997
ecr 40, True
ecr 2450450#

ecrVariable "$TDUPDATE"
ecr 40, True
ecr 2450450#

ecrVariable "$TDINDWG"
ecr 40, True
ecr 0#

ecrVariable "$TDUSRTIMER"
ecr 40, True
ecr 0#

ecrVariable "$USRTIMER"
ecr 70, True
ecr 1

ecrVariable "$ANGBASE"
ecr 50, True
ecr 0#

ecrVariable "$ANGDIR"
ecr 70, True
ecr 0

ecrVariable "$PDMODE"
ecr 70, True
ecr 0

ecrVariable "$PDSIZE"
ecr 40, True
ecr 0#

ecrVariable "$PLINEWID"
ecr 40, True
ecr 0#

ecrVariable "$SPLFRAME"
ecr 70, True
ecr 0

ecrVariable "$SPLINETYPE"
ecr 70, True
ecr 6

ecrVariable "$SPLINESEGS"
ecr 70, True
ecr 8

' Maintiens

If VersionDXF = 11 Then
  ecrVariable "$HANDLING"     ' Maintiens :
  ecr 70, True                'ignorés en v12(pour GIRATION)- valeur=0
  ecr 0                       ' deviendront obligatoires en v14(et du coup la variable devient inutile : forcément 1)
  
  ecrVariable "$HANDSEED"
  ecr 5, True
  ecr "0"
  
Else
  ecrVariable "$HANDSEED"
  ecr 5, True
  CalculerMaxHandle
  ecr MaxHandle
  IDAcad = 0
End If

ecrVariable "$SURFTAB1"
ecr 70, True
ecr 6

ecrVariable "$SURFTAB2"
ecr 70, True
ecr 6

ecrVariable "$SURFTYPE"
ecr 70, True
ecr 6

ecrVariable "$SURFU"
ecr 70, True
ecr 6

ecrVariable "$SURFV"
ecr 70, True
ecr 6

ecrVariable "$UCSNAME"
ecr 2, True
ecr ""

ecrVariable "$UCSORG"
ecr 10, True
ecr 0#
ecr 20, True
ecr 0#
ecr 30, True
ecr 0#

ecrVariable "$UCSXDIR"
ecr 10, True
ecr 1#
ecr 20, True
ecr 0#
ecr 30, True
ecr 0#

ecrVariable "$UCSYDIR"
ecr 10, True
ecr 0#
ecr 20, True
ecr 1#
ecr 30, True
ecr 0#

ecrVariable "$PUCSNAME"
ecr 2, True
ecr ""

ecrVariable "$PUCSORG"
ecr 10, True
ecr 0#
ecr 20, True
ecr 0#
ecr 30, True
ecr 0#

ecrVariable "$PUCSXDIR"
ecr 10, True
ecr 1#
ecr 20, True
ecr 0#
ecr 30, True
ecr 0#

ecrVariable "$PUCSYDIR"
ecr 10, True
ecr 0#
ecr 20, True
ecr 1#
ecr 30, True
ecr 0#

ecrVariable "$USERI1"
ecr 70, True
ecr 0

ecrVariable "$USERI2"
ecr 70, True
ecr 0

ecrVariable "$USERI3"
ecr 70, True
ecr 0

ecrVariable "$USERI4"
ecr 70, True
ecr 0

ecrVariable "$USERI5"
ecr 70, True
ecr 0

ecrVariable "$USERR1"
ecr 40, True
ecr 0#

ecrVariable "$USERR2"
ecr 40, True
ecr 0#

ecrVariable "$USERR3"
ecr 40, True
ecr 0#

ecrVariable "$USERR4"
ecr 40, True
ecr 0#

ecrVariable "$USERR5"
ecr 40, True
ecr 0#

ecrVariable "$WORLDVIEW"
ecr 70, True
ecr 1

ecrVariable "$SHADEDGE"
ecr 70, True
ecr 3

ecrVariable "$SHADEDIF"
ecr 70, True
ecr 70

ecrVariable "$TILEMODE"
ecr 70, True
ecr 1

ecrVariable "$MAXACTVP"
ecr 70, True
ecr 16

ecrVariable "$PINSBASE"
ecr 10, True
ecr 0#
ecr 20, True
ecr 0#
ecr 30, True
ecr 0#

ecrVariable "$PLIMCHECK"
ecr 70, True
ecr 0

ecrVariable "$PEXTMIN"       ' Extensions de l'espace papier
ecr 10, True
ecr 1E+20
ecr 20, True
ecr 1E+20
ecr 30, True
ecr 1E+20

ecrVariable "$PEXTMAX"
ecr 10, True
ecr -1E+20
ecr 20, True
ecr -1E+20
ecr 30, True
ecr -1E+20

ecrVariable "$PLIMMIN"
ecr 10, True
ecr 0#
ecr 20, True
ecr 0#

ecrVariable "$PLIMMAX"
ecr 10, True
ecr 12#
ecr 20, True
ecr 9#

ecrVariable "$UNITMODE"
ecr 70, True
ecr 0

ecrVariable "$VISRETAIN"
ecr 70, True
ecr 0

ecrVariable "$PLINEGEN"
ecr 70, True
ecr 0

ecrVariable "$PSLTSCALE"
ecr 70, True
ecr 0

ecrVariable "$TREEDEPTH"
ecr 70, True
ecr 3020

End Sub

Private Sub CalculerMaxHandle()
Dim nBloc As BLOC
Dim nPline As POLYLINE
Dim i As Long

  'Décompte des 20 entrées minimales
  ' 9 Entrées de tables générales (LAYER, LTYPE ...)
  ' 1 Table VPORT : *Active
  ' 3 Tables LTYPE : BYLAYER - BYBLOCK - CONTINUOUS
  ' 1 Table STYLE : STANDARD
  ' 1 Table APPID : ACAD
  ' 3 Entrées de dictionnaire
  i = 18
  
  'Une entrée par plan et 3 entrées par bloc (BLOCK_RECORD - BLOCK - ENDBLK)
  ' En plus des blocs : 2 blocs *MODEL_SPACE et * PAPER_SPACE, qui seront rajoutés dans ecrTableBloc
  i = i + nbPlans + (mBlocs.Count + 2) * 3

  For Each nBloc In mBlocs
    With nBloc
      i = i + .blkPoints.Count
      i = i + .blkLines.Count
      i = i + .blkPolylines.Count * 2
      For Each nPline In .blkPolylines
        i = i + nPline.Pts.Count
      Next
      i = i + .blkSplines.Count
      i = i + .blkCercles.Count
      i = i + .blkArcs.Count
      i = i + .blkInserts.Count
    End With
  Next
  
  If EcrireFDP Then
    With fCourante.GraphFDP
      i = i + .gphPoints.Count
      i = i + .gphLines.Count
      i = i + .gphPolylines.Count * 2
      For Each nPline In .gphPolylines
        i = i + nPline.Pts.Count
      Next
      i = i + .gphSplines.Count
      i = i + .gphCercles.Count
      i = i + .gphArcs.Count
      i = i + .gphInserts.Count
    End With
  
  End If

  If flagExport Then
    ' Fonction Export
    With fCourante.GraphExport
      i = i + .gphPoints.Count
      i = i + .gphLines.Count
      i = i + .gphPolylines.Count * 2
      For Each nPline In .gphPolylines
        i = i + nPline.Pts.Count
      Next
      i = i + .gphSplines.Count
      i = i + .gphCercles.Count
      i = i + .gphArcs.Count
      i = i + .gphInserts.Count
    End With

  End If
  
  MaxHandle = Hex(i + 1)
  
End Sub

Private Sub ecrireTables()

  ecrTableViewport
  ecrTableLtype
  ecrTableLayer
  ecrTableStyle
  ecrTablesInutilisées
  If VersionDXF > 11 Then ecrTableBloc

End Sub

Private Sub ecrireDictionnaires()
Dim strID(0 To 1) As String

  ecrArticle "DICTIONARY", AvecID:=True
  ecr 330, SansCadrage:=True
  ecr 0
  ecr 100, SansCadrage:=True
  ecr "AcDbDictionary"
  
  ecr 3, SansCadrage:=True
  ecr "ACAD_GROUP"
  strID(0) = ecrNewID(350)
  
  ecr 3, SansCadrage:=True
  ecr "ACAD_MLINESTYLE"
  strID(1) = ecrNewID(350)
  
  ecrArticle "DICTIONARY"
  ecr 5, SansCadrage:=True
  ecr strID(0)
  ecr 100, SansCadrage:=True
  ecr "AcDbDictionary"
  
  ecrArticle "DICTIONARY"
  ecr 5, SansCadrage:=True
  ecr strID(1)
  ecr 100, SansCadrage:=True
  ecr "AcDbDictionary"
  ecr 3, SansCadrage:=True
  ecr "STANDARD"
  ecrNewID (350)
  
End Sub

Public Sub ecr(ByVal v As Variant, Optional ByVal SansCadrage As Boolean)
Dim s$, pos%

  Select Case VarType(v)
  Case vbInteger, vbLong
    s = Mid(str(v), 2)

    If Not SansCadrage Then
      s = "  " & s
      If s < 1000 Then s = " " & s
    End If

    If v < 10 Then
      s = "  " & s
    ElseIf v < 100 Then
      s = " " & s
    End If

  Case vbSingle, vbDouble
    pos = InStr(1, str(v), "E", 1)
    If pos <> 0 Then
      s = Format(v, "0.000000E+00")
    Else
      s = Format(v, "0.0#########")
    End If
    substPointVirgule s
  
  Case vbString
    s = v
  End Select
  
  s = s & Chr(13) & Chr(10)
  Put #numFich, , s
  
End Sub

Private Sub substPointVirgule(ByRef s As String)
  Dim pos%
  
    pos = InStr(s, Chr(gtPtDecimal))
    If pos <> 0 Then Mid(s, pos) = "."
End Sub

Public Sub ecrArticle(ByVal Nom As String, Optional ByVal AvecID As Boolean)
  ' Ecriture d'un nouvel article ("SECTION","TABLE","POLYLINE","ENDSEC" ,,,,)
  ecr 0, SansCadrage:=True
  ecr Nom
  
  If AvecID Then ecrNewID 5
  
End Sub

Private Function ecrNewID(ByVal CodeID As Integer) As String
  
  ecrNewID = IncrémenterID
  
  ecr CodeID, SansCadrage:=True
  ecr ecrNewID
  
End Function

Private Function IncrémenterID() As String

  IDAcad = IDAcad + 1
  IncrémenterID = Hex(IDAcad)

End Function

Private Function FormatNomV14(ByVal chaine As String) As String
Dim pos As Integer
Dim chaine2 As String

  'Dans AutoCAD 2007 les noms sont susceptibles de comporter des blancs et de ne pas être en majuscules
  'Peut-être aussi à partir d'Autocad 2004
  
  chaine = UCase(chaine)
  
  pos = InStr(chaine, " ")
  If pos = 0 Then
    FormatNomV14 = chaine
    
  Else
    Do
    'Remplacer l'espace par un  '_'
      chaine2 = chaine2 & Left(chaine, pos - 1) & "_"
      chaine = Mid(chaine, pos + 1)
      pos = InStr(chaine, " ")
    Loop Until pos = 0
    FormatNomV14 = chaine2 & chaine
  End If

End Function
  
Public Sub ecrNom(ByVal Nom As String)
  ' Ecriture d'un nom
  ecr 2, True
  ecr FormatNomV14(Nom)
  
End Sub

Private Sub ecrVariable(ByVal Nom As String)
' Ecriture d'une variable dans la section HEADER
  ecr 9, True
  ecr Nom
  
End Sub

Public Sub ecrPlan(ByVal Nom As String)
  ' Ecriture du nom de plan d'une entité
  ecr 8, True
  ecr Nom
  
End Sub

Public Sub ecrCouleur(ByVal couleur As Long)
  ' Ecriture de la couleur d'une entité
  
  If couleur <> 256 Then   ' On ne renseigne pas si DUPLAN
    ecr 62, True
    ecr couleur
  End If

End Sub

Public Sub ecrTypelign(ByVal typelign As String)
  ' Ecriture du type de ligne d'une entité

  If typelign <> "BYLAYER" Then   ' On ne renseigne pas si DUPLAN
    ecr 6, True
    If typelign <> "CONTINUOUS" And typelign <> "BYBLOCK" Then
'      ecr "POINTILLE"
      ecr "CONTINUOUS"    ' debug : si on incorpore à nouveau pointille, réactiver la ligne ci-dessus
    Else
      ecr typelign
    End If
  End If

End Sub

Public Sub ecrPoint(ByVal p As PT, Optional ByVal flag As Variant)
Dim increment%
' la plupart des codes DXF d'un point sont 10,20,30
' pour le 2ème point d'une ligne, on rajoute 1 pour obtenir 11,21,31

If Not IsMissing(flag) Then increment = flag
  ecr 10 + increment, True
  ecr p.x
  ecr 20 + increment, True
  ecr p.y
'  ecr 30 + increment, True
'  ecr 0#
  
End Sub

Private Function acadCol(couleur As Long) As Long
Dim i%

  For i = 0 To 15
    If couleur = QBColor(i) Then
      acadCol = InvTCouleur(i)
      Exit For
    End If
  Next
  
End Function
