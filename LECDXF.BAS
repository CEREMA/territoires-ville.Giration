Attribute VB_Name = "LecDXF"
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'         Mise à jour : Octobre 1999 (DXF14 et 2000)
'                       Février 2000 (Simplification de lire_entité pour les polylignes fermées)
'                       Novembre 2000(Simplisfication de lire_chaine)
'   Réalisation : André VIGNAUD

'   Module standard : LecDXF   -   LECDXF.BAS
'   Lecture d'un fichier DXF

'**************************************************************************************
Option Explicit

Private Binaire As Boolean
Private CODE As Integer, chaine As String
Private NomFich As String, CodeGroupe As String, ValeurGroupe As String, CodeGroupe2 As String, ValeurGroupe2 As String, CodeGroupe3 As String, ValeurGroupe3 As String
Private num_ligne As Long
Private numFich As Integer
Private nomCourtVéhicule As String
Private IndiceVéhicule() As Integer

Private UnPourCent As Single
Private PourCentEnCours As Integer

Public Graph As GRAPHIQUE
Public echDXF As Double, Abandon As Double

Private dwgVersion As Integer
 
Public Enum TypeDonnBinaire
  typeChaine
  typeCoord
  typeDouble
  typeEtendueBinaire
  typeEntier
  typeLong
  typeBooléen
  
End Enum

'**************************************************************************************
' Lecture dun fichier Fond de plan au format DXF
' FDP : nom du Fichier
' Cancel : Retourne Vrai en cas d'erreur
'**************************************************************************************
 Public Sub lecFDP(ByVal FDP As String, ByRef Cancel As Boolean)
Dim planNew As New PLAN
Dim i%, j%
Dim PremièreLigne As String
  
  Binaire = False
  NomFich = FDP
  echDXF = 1
  dwgVersion = 12
  
  Set gtPlans = fCourante.colPlans        ' gtPlans est déclaré dans GRAPHST.BAS
  Set Graph = fCourante.GraphFDP
  nomCourtVéhicule = fCourante.Vehic.nomCourt
  ReDim IndiceVéhicule(fCourante.nbRemorque)
  For i = 0 To fCourante.nbRemorque
    IndiceVéhicule(i) = -1
  Next
  
      ' Cette réinitialisation n'aua pas lieu d'être si on souhaite la superposition de FDP
  reinFDP
  
  On Error GoTo GestErr
  
  fCourante.pminFDP.x = 0
  fCourante.pminFDP.y = 0
  fCourante.pmaxFDP.x = 0
  fCourante.pmaxFDP.y = 0

'============= Panneau de défilement  =================
  With frmDéfileDXF
    .Show
    .tmrDéfile.Enabled = True
    .txtPanneau = ID_LectureFichier & " " & Trim(NomFich) & "..." & ID_LectureFichier & Trim(NomFich) & "..."
    UnPourCent = FileLen(NomFich) / 100
    PourCentEnCours = 0
    .lblPourCent = "0%"
    .SetFocus
  End With
  fCourante.MousePointer = vbHourglass  'Sablier


  CalcDuree True
  
'============= Détermination du type ASCII/Binaire =================
  num_ligne = -1
  numFich = FreeFile
  Open NomFich For Input As numFich
  Input #numFich, PremièreLigne
  If PremièreLigne = "AutoCAD Binary DXF" Then
    Binaire = True
  End If
  Close numFich
    
  If Binaire Then
    Dim octet As Byte
    Open NomFich For Binary As numFich
    Get #numFich, 24, octet
    If octet = 0 Then dwgVersion = 14
    Seek #numFich, 23
  Else
    Open NomFich For Input As numFich ' Len = 32767 debug: cette option n'améliore en rien la rapidité de lecture
  End If
  
  
  lire_code 0
  lire_chaine "SECTION"
  lire_code 2
'        lire_chaine "HEADER", "TABLES", "ENTITIES"  Modif : v14
  lire_chaine "HEADER", "CLASSES", "TABLES", "ENTITIES"
  
'============= Section HEADER =================
  If chaine = "HEADER" Then ' Certains fichiers DXF n'ont pas de section HEADER - c'est bizarre, mais AutoCAD l'accepte
    lire_code 9
    While (CODE = 9)
      lire_header
    Wend
    lire_chaine "ENDSEC"
    lire_code 0
    lire_chaine "SECTION"
    lire_code 2
'          lire_chaine "TABLES"    Modif : v14
    lire_chaine "CLASSES", "TABLES", "ENTITIES"
  End If                     ' chaine = "HEADER"
  
'============= Section CLASSES =================
  If chaine = "CLASSES" Then        ' Fichier DXF  v14 : section CLASSES
    lire_code 0
    lire_chaine "CLASS", "ENDSEC"
    While chaine = "CLASS"
      lire_classe
      lire_chaine "CLASS", "ENDSEC"
    Wend
    lire_code 0
    lire_chaine "SECTION"
    lire_code 2
    lire_chaine "TABLES", "ENTITIES"
  End If
  
'============= Section TABLES =================
  If chaine = "TABLES" Then     ' Sinon ce fichier DXF n'a qu'1 section ENTITIES
    lire_code 0
    lire_chaine "TABLE", "ENDSEC"
    While chaine = "TABLE"
      lire_table
      lire_chaine "TABLE", "ENDSEC"
    Wend
  
' Au cas, peu probable, où le plan 0 serait absent de la section
    planNew.name = "0"
    AjouterPlan planNew
           
    If Extension(FDP, Conversion:=vbLowerCase) <> "fdp" Then
      ' Choix des plans à conserver
      With frmDéfileDXF
        .tmrDéfile.Enabled = False
        frmImport.Show vbModal    ' 1 feuille modale
        If Graph Is Nothing Then Err.Raise 102
        .tmrDéfile.Enabled = True
        .SetFocus
      End With
      fCourante.pminFDP.x = fCourante.pminFDP.x * echDXF
      fCourante.pminFDP.y = fCourante.pminFDP.y * echDXF
      fCourante.pmaxFDP.x = fCourante.pmaxFDP.x * echDXF
      fCourante.pmaxFDP.y = fCourante.pmaxFDP.y * echDXF
    End If
  
    lire_code 0
    lire_chaine "SECTION"
    lire_code 2
    lire_chaine "BLOCKS", "ENTITIES"
    
'============= Section BLOCKS =================
    If chaine = "BLOCKS" Then ' Certains fichiers DXF (PISTE,TPL) n'ont pas de section BLOCKS - c'est bizarre, mais AutoCAD l'accepte
      lire_code 0
      lire_chaine "BLOCK", "ENDSEC"
      While chaine = "BLOCK"
          lire_bloc
          lire_chaine "BLOCK", "ENDSEC"
      Wend
      MajListeBlocs
      lire_code 0
      lire_chaine "SECTION"
      lire_code 2
      lire_chaine "ENTITIES"
    End If                                        ' chaine = "BLOCKS"

  Else  ' Le fichier ne contient qu'1 section ENTITIES
    planNew.name = "0"
    AjouterPlan planNew
  End If                                          ' chaine =  "TABLES"

'============= Section ENTITIES =================
  lire_code 0
  lire_chaine "", "ENDSEC"
  Do While chaine <> "ENDSEC"
    lire_entite chaine
    lire_chaine "", "ENDSEC"
  Loop

  lire_code 0
'        lire_chaine "EOF"         ' modif V14
  lire_chaine "SECTION", "EOF"
  
'============= Section OBJECTS =================
  If chaine = "SECTION" Then
    lire_code 2
    lire_chaine "OBJECTS"
    lire_code 0
    lire_chaine "", "ENDSEC"
    
    'Mise en commentaire (AV : 25/06/03) des lignes qui suivent
    'GIRATION n'utilise pas la section OBJECTS : inutile de les lire
    
'    While chaine <> "ENDSEC"
'      lire_objet
'      lire_chaine "", "ENDSEC"
'    Wend
'
''============= Section EOF =================
'    lire_code 0
'    lire_chaine "EOF"

  End If
  Close numFich
  

'        CalcDuree False
  frmDéfileDXF.tmrDéfile.Enabled = False
  frmDéfileDXF.cmdCancel.Enabled = False
'        DoEvents
  
'=============   ' traitement final pour le calcul de Zoom =================
  With fCourante
    If .pminFDP.x = 0 And .pminFDP.y = 0 And .pmaxFDP.x = 0 And .pmaxFDP.y = 0 Then
    ' Pas de section HEADER
      frmDéfileDXF.txtPanneau = ID_RechercheLimites
      Wpmin.x = 1000000000#
      Wpmin.y = 1000000000#
      Wpmax.x = -1000000000#
      Wpmax.y = -1000000000#
      calminmax
      .pminFDP.x = Wpmin.x
      .pminFDP.y = Wpmin.y
      .pmaxFDP.x = Wpmax.x
      .pmaxFDP.y = Wpmax.y
    Else
      Wpmin.x = .pminFDP.x
      Wpmin.y = .pminFDP.y
      Wpmax.x = .pmaxFDP.x
      Wpmax.y = .pmaxFDP.y
    End If
  End With    ' Fin du traitement final
    

  fCourante.MousePointer = vbDefault
  Unload frmDéfileDXF

'        If fCourante.pmin.X = 10000000000# Then Err.Raise 101


  Exit Sub

'============= Traitement d'erreur =================
GestErr:
    Close numFich
    Unload frmDéfileDXF
    ' Err=100 est déclenché par le programme dans FinRapide (Une erreur + explicite a été fournie à l'utilisateur)
    ' Err=102 si l'utilisateur a fait Escape dans frmDéfileDXF ou Annuler dans la frmImport
    If Err <> 100 And Err <> 102 Then
      ErreurFDP CStr(Err) & " " & Err.Description
    End If
    fCourante.MousePointer = vbDefault
    fCourante.pminFDP.x = 0
    fCourante.pminFDP.y = 0
    fCourante.pmaxFDP.x = 0
    fCourante.pmaxFDP.y = 0
    Cancel = True
    reinFDP         '  idem
        
End Sub

'*******************************************
'* lire_objet : lecture d'un objet si DXF 14
'*******************************************
' Ne fait rien

Private Sub lire_objet()

  lire_code -1
  
  While CODE <> 0
    lire_chaine ""
    lire_code -1
  Wend
  
End Sub

'*******************************************
'* lire_clsse : lecture d'une classe si DXF 14
'*******************************************
' Ne fait rien

Private Sub lire_classe()

  lire_code 1
  While CODE <> 0
    lire_chaine ""
    lire_code -1
  Wend

End Sub


Private Sub ErreurFDP(ByVal mess As String, Optional ByVal NonBloquant As Boolean)

Dim mess_err As String

If Err = 101 Then
  mess_err = IDm_Err101 & " " & NomFich
ElseIf Err = 103 Then
  mess_err = IDm_Err103
Else
  If Binaire Then
    mess_err = mess & vbCrLf & ID_Fichier & " " & NomFich
  Else
    mess_err = mess & vbCrLf & ID_LIGNE & " " & CStr(num_ligne) & " - " & ID_Fichier & " " & NomFich
  End If
  
  If NonBloquant Then
    mess_err = IDm_Erreur & " : " & mess_err
  Else
    mess_err = IDm_Anomalie & " : " & mess_err
  End If
End If

If NonBloquant Then
  MsgBox mess_err, vbExclamation + vbOKOnly, ID_LectureFichier & "DXF"
Else
  mess_err = mess_err & Chr(13) & IDm_FDPRefusé
  MsgBox mess_err, vbCritical + vbOKOnly, ID_LectureFichier & "DXF"
  If Err = 0 Then Err.Raise 100
End If

End Sub

Private Sub lire_header()
Dim flag(1) As Boolean

    lire_chaine ""   '               // nom de la variable AutoCad
    lire_code -1                      '               // lecture du 1er groupe pour la variable
    
    Select Case chaine
    Case "$ACADVER"
      Select Case ValeurGroupe
        Case "AC1009"
          dwgVersion = 12 ' v11 aussi
        Case "AC1012"
          dwgVersion = 13
        Case "AC1014"
          dwgVersion = 14 ' bien que la doc de la v14 dise "AC1013" - mais c'est a priori "AC1014" (cf Aide Autocad 2000)
        Case "AC1015"
          dwgVersion = 15 ' AutoCAD 2000,200i(16) et 2002(17)
        Case "AC1018"
          dwgVersion = 18  ' Autocad 2004,5(19),6(20)
          'Ajout version 3.3.1
        Case "AC1021"
          dwgVersion = 21  ' Autocad 2007
        Case "AC1006"
          ErreurFDP ID_DXFVersion & CStr(10) & ID_NonGeree & App.Title
        Case Else
          ErreurFDP ID_DXFVersion & ID_NonGeree & App.Title
      End Select
      lire_code -1

    ' Recherche des limites utiles du dessin
    Case "$EXTMIN"
     While Not (flag(0) And flag(1))
      If CODE = 10 Then
        If val(ValeurGroupe) <> 1E+20 Then     ' version Site pilote
          fCourante.pminFDP.x = val(ValeurGroupe)
        End If
        flag(0) = True
      ElseIf CODE = 20 Then
        If val(ValeurGroupe) <> 1E+20 Then     ' version Site pilote
          fCourante.pminFDP.y = val(ValeurGroupe)
        End If
        flag(1) = True
      End If
      lire_code -1                    '                 lecture du groupe suivant
    Wend
    
  Case "$EXTMAX"
    While Not (flag(0) And flag(1))
      If CODE = 10 Then
        If val(ValeurGroupe) <> -1E+20 Then     ' version Site pilote
          fCourante.pmaxFDP.x = val(ValeurGroupe)
        End If
        flag(0) = True
      ElseIf CODE = 20 Then
        If val(ValeurGroupe) <> -1E+20 Then     ' version Site pilote
          fCourante.pmaxFDP.y = val(ValeurGroupe)
        End If
        flag(1) = True
      End If
      lire_code -1                    '                 lecture du groupe suivant
    Wend
    
  End Select
   
  While (CODE <> 9 And CODE <> 0)
    lire_code -1
  Wend

End Sub
        
Private Sub lire_table()
  Dim GroupeTable As String, nomTable As String
  Dim i As Integer
'    dim nbTables As Integer
  Dim planNew As PLAN
        
    lire_code 2
    lire_chaine "" ' // nom du groupe de tables
    GroupeTable = chaine
    '<===AV : 14/01/2004
        'If dwgVersion <> 12 Then ..........
    'A partir de la v3.2.0032(01/2004), la gestion des codes spécifiques v14,2000... est déportée dans lire_code
    'Le dernier source est disponible dans LexDXFv32.Bas (sous vSpline)
    '===>AV
    lire_code 70, 0
    If CODE = 70 Then
      If dwgVersion >= 15 And GroupeTable = "DIMSTYLE" Then
        lire_code 71
      End If
      lire_code 0
    End If
      
    Do While True
      lire_chaine GroupeTable, "ENDTAB"
      If chaine = "ENDTAB" Then   ' moins de tables que nbTables
        lire_code 0
        Exit Do
      Else
        lire_code 2
        lire_chaine ""    '       // nom utilisateur de la table
        nomTable = chaine
        If GroupeTable = "LAYER" Then
          Set planNew = New PLAN
          planNew.name = nomTable
        End If
        lire_code -1
        While (CODE <> 0)
            If GroupeTable = "LAYER" Then
              If CODE = 70 Then
                planNew.gele = planNew.gele Or BitActif(val(ValeurGroupe), 1)
              ElseIf CODE = 62 Then
                planNew.couleur = val(ValeurGroupe)
                planNew.gele = planNew.gele Or (val(ValeurGroupe) < 0)   ' plan inactif dans le DXF
              ElseIf CODE = 6 Then
                planNew.typelign = ValeurGroupe
              End If
            End If
            lire_code -1
        Wend
        If GroupeTable = "LAYER" Then AjouterPlan planNew
      End If
    Loop
'        Next i
    
    If chaine <> "ENDTAB" Then    ' Il n'y a pas eu de sortie anticipée dans la boucle de lecture
      lire_chaine "ENDTAB"
      lire_code 0
    End If
        
End Sub
        
Private Sub AjouterPlan(ByVal objPlan As PLAN)
    
  'Ajout AV (19/06/03 :suite à pb sous ACAD2000   (ECRDXF) : Points de définition des cotes de toute façon ignoré de GIRATION
  If objPlan.name = "DEFPOINTS" Then Exit Sub
  
'Ajout AV (19/06/03) : en v14, la lecture d'un DXF12 avec une couleur 0 est refusée par AutoCAD
  If objPlan.couleur = 0 Then objPlan.couleur = 7
  
  If objPlan.typelign = "" Then objPlan.typelign = "CONTINUOUS"
  
  If Not gtPlans.Existe(objPlan.name) Then gtPlans.Add objPlan

End Sub

Private Sub lire_bloc()
Dim nomBlocUser As String, p As New PT
Dim BlocOK As Boolean
Dim Valeur As Integer
Dim objetBloc As BLOC

  lire_code -1
  While (CODE <> 0)   '            // en-tete du bloc jusqu'à la première entité
    Select Case CODE
      Case 3, 8, 62, 30
      Case 2:
        nomBlocUser = ValeurGroupe
      Case 10
        p.x = val(ValeurGroupe)
      Case 20
        p.y = val(ValeurGroupe)
        
      Case 70:
        Valeur = val(ValeurGroupe)
        If Valeur = 64 Or Valeur = 66 Then
          BlocOK = True
        ElseIf Valeur = 0 Or Valeur = 2 Then
        ' Ajout AV : 07/01/2002 - ces blocs sont des blocs réservés AutoCAD pour l'espace papier et l'espace objet
          If InStr(1, nomBlocUser, "*Paper_Space", vbTextCompare) = 1 Or _
                UCase(nomBlocUser) = "*MODEL_SPACE" Then
            BlocOK = False
          Else  ' curieusement, les fichiers MOSS mettent un indicateur 0 au lieu de 64 pour les blocs référencés et AutoCAD les lit quand même
            BlocOK = True
          End If
        Else  '         // bloc non classique (xref, ou dépendant d'xref, ou anonyme) ou non référencé
          BlocOK = False
        End If
        
        If Not BlocOK Then
          Do
            lire_code -1
            lire_chaine ""
          Loop While chaine <> "ENDBLK"
        End If
      End Select
      
      lire_code -1
    Wend
  
  If BlocOK Then                        '         // bloc non ignoré
    ' vbTextCompare ne fait pas de différence Majuscule/Minuscule
    If InStr(1, nomBlocUser, "GIR_" & nomCourtVéhicule, vbTextCompare) = 1 Then
      VerifNomblocReserve nomBlocUser, Creation:=True
    End If
    Set objetBloc = Graph.gphBlocs.Add(nomBlocUser)
    
'Ajouter toutes les entités composant le bloc
    lire_chaine "", "ENDBLK"
    While chaine <> "ENDBLK"
      lire_entite chaine, p, objetBloc
      lire_chaine "", "ENDSEC"
    Wend
    lire_code -1
    
    If objetBloc.nbElements = 0 Then
      Graph.gphBlocs.Remove nomBlocUser
    End If
  End If
  
  While (CODE <> 0)         '     // contrairt à la doc d'AutoDESK, ENDBLK peut avoir des codes de groupe (plan, maintien....autres???)
    lire_chaine ""
    lire_code -1
  Wend
End Sub

'*******************************************************************************
' A la fin de la création des blocs, on vérifie au cas où le véhicule comporte
' plusieurs éléments de carrosserie (lire commentaire ci-après pour VerifNomblocReserve)
' à ce que les blocs se terminent par le même $n
'*******************************************************************************
Private Sub MajListeBlocs()
Dim i As Integer, j As Integer
Dim NomReduit As String, numRemorque As Integer

  For numRemorque = 1 To fCourante.nbRemorque
    j = Max(IndiceVéhicule(numRemorque), IndiceVéhicule(numRemorque - 1))
  Next
  
  If fCourante.nbRemorque = 0 Or j = -1 Then
    ' Pas de bloc réservé
  Else
    NomReduit = Left(nomCourtVéhicule, Min(LGMAXNOMVEHICULE - 2, Len(nomCourtVéhicule)))
    For numRemorque = 0 To fCourante.nbRemorque
      i = IndiceVéhicule(numRemorque)
      If i < j And i <> -1 Then
        Graph.gphBlocs.Renommer _
          "GIR_" & NomReduit & "$" & CStr(i) & radicExport(numRemorque), _
          "GIR_" & NomReduit & "$" & CStr(j) & radicExport(numRemorque)
        IndiceVéhicule(numRemorque) = j
      End If
    Next
  End If

End Sub

'******************************************************************************
' Si on importe un DXF comportant déjà un nom de bloc issu de l'export (GIR_<nomvehic>_<suffixecarrosserie>)
' il ne faut pas qu'un deuxième export crée un bloc de même nom (il faut pour cela que ce soit le même véhicule)
' On substitue par précaution $0 aux 2 derniers caractères du véhicule
' Et par super-précaution, $1,$2... si l'utilisateur refait la manip plusieurs fois
'******************************************************************************
Private Sub VerifNomblocReserve(ByRef nomBlocUser As String, ByVal Creation As Boolean)
Dim Nom As String, NomReduit As String, numRemorque As Integer
Dim i As Integer

  ' Recherche de l'élément de carrosserie
  Select Case Mid(UCase(nomBlocUser), 5 + Len(nomCourtVéhicule))
  Case UCase(radicExport(TRACTEUR))
    numRemorque = TRACTEUR
  Case UCase(radicExport(REMORQUE1))
    numRemorque = REMORQUE1
  Case UCase(radicExport(REMORQUE2))
    numRemorque = REMORQUE2
  End Select
  
  ' Le cas peut arriver si le nom du véhicule est < 14 car.  : on conserve les blocs ayant déjà un '$n' dans le nom
  If InStr(1, nomBlocUser, "GIR_" & nomCourtVéhicule & radicExport(numRemorque), vbTextCompare) = 1 Then
  
    NomReduit = Left(nomCourtVéhicule, Min(LGMAXNOMVEHICULE - 2, Len(nomCourtVéhicule)))
    If Creation Then
      Nom = "GIR_" & NomReduit & "$" & CStr(i) & radicExport(numRemorque)
      On Error GoTo GestErr
      Do While TypeOf Graph.gphBlocs(Nom) Is BLOC
        i = i + 1
        Nom = "GIR_" & NomReduit & "$" & CStr(i) & radicExport(numRemorque)
      Loop
    Else  ' Insertion des objets : récupération du nom à partir du numéro de carrosserie
      nomBlocUser = "GIR_" & NomReduit & "$" & CStr(IndiceVéhicule(numRemorque)) & radicExport(numRemorque)
    End If
  End If

Exit Sub
  
GestErr:
  If Err = 5 Then
    nomBlocUser = Nom
    IndiceVéhicule(numRemorque) = i ' Mémorisation, en vue de l'insert ultérieur (lire_entite)
  Else
    ErrGeneral
  End If

End Sub

'******************************************************************************
' Lecture d'une entité élémentaire ou incluse dans un bloc
'******************************************************************************
Private Sub lire_entite(ByVal nomEntite As String, Optional ByVal pBloc As PT, Optional ByVal objetBloc As BLOC)
Dim interessant As Boolean
Dim nomPlan As String, couleur As Long, CodeBinaire As Integer
Dim p0 As New PT
Dim p1 As New PT
Dim nomBloc As String, typelign As String
Dim echx As Double, echy As Double, rot As Double, nbcol As Integer, nbrow As Integer, distcol As Double, distrow As Double
Dim Rayon As Double, Alpha(1) As Double
Dim arrondi() As Double
Dim Drapeau As Integer
Dim Sommets As Pts
Dim pLoop As PT
Dim planNew As New PLAN
Dim numVertex As Integer, nbVertex As Integer
' SPLINE
Dim nbNoeuds As Integer, numNoeud As Integer
Dim mSpLine As SPLINE
Dim mPolylines As Polylines

Static pLine As POLYLINE

Static flagArrondi As Boolean

On Error GoTo GestErr

  If pBloc Is Nothing Then Set pBloc = New PT
  
  Select Case nomEntite
  Case "POINT", "LINE", "ARC", "CIRCLE", "INSERT"
    interessant = True
    If nomEntite = "INSERT" Then
      echx = 1: echy = 1: nbcol = 1: nbrow = 1
    End If
  Case "POLYLINE", "LWPOLYLINE"
    interessant = True
    Set pLine = New POLYLINE
  Case "VERTEX", "SEQEND"
    interessant = Not pLine Is Nothing
    ReDim arrondi(0)
  Case "SPLINE"
    interessant = True
    Set mSpLine = New SPLINE
  End Select
  
  couleur = 256                   'par défaut: couleur DUPLAN
  typelign = "BYLAYER"      'par défaut: type de ligne DUPLAN
  
  lire_code -1
  
  Do While (CODE <> 0)
    If interessant Then
      
      Select Case CODE
      Case 30, 31, 66  ' infos ignorées même pour les entités intéressantes
      Case 43
        If nomEntite = "SPLINE" Then mSpLine.tolPControle = val(ValeurGroupe)
      Case Is > 997   ' Données étendues ignorées
      Case 8
        nomPlan = ValeurGroupe
        If Not gtPlans.Existe(nomPlan) Then
          If nomPlan = "DEFPOINTS" Then
  'Ajout AV (19/06/03 :suite à pb sous ACAD2000   (ECRDXF) : Points de définition des cotes de toute façon ignoré de GIRATION
            interessant = False
          Else
            ErreurFDP ID_Plan & " " & nomPlan & IDm_AbsentTablePlan, NonBloquant:=True
            ' AutoCAD crée le plan s'il est absent : GIRATION aussi
            planNew.name = nomPlan
            AjouterPlan planNew
          End If
        Else
          If gtPlans(nomPlan).gele Then interessant = False
        End If
      Case 62
        couleur = val(ValeurGroupe)
      Case 6
        typelign = ValeurGroupe
      Case 70 'insert,polyline,vertex
        Select Case nomEntite
        Case "INSERT"
          nbcol = val(ValeurGroupe)
        Case "LWPOLYLINE"
          CodeBinaire = val(ValeurGroupe)
          'pLine.ferme = ((CodeBinaire And 1) = 1)
          pLine.ferme = BitActif(CodeBinaire, 1)
        Case "POLYLINE"
          CodeBinaire = val(ValeurGroupe)
          If CodeExcluPoly(CodeBinaire) Then
            interessant = False
          Else
            pLine.ferme = BitActif(CodeBinaire, 1)
          End If
        Case "VERTEX"
          CodeBinaire = val(ValeurGroupe)
          If CodeExcluSommet(CodeBinaire) Then interessant = False
          'Oter le Drapeau P3D pour ne pas le réécrire lors de l'export (GIRATION lit mais ignore le Z du 3D)
          If BitActif(CodeBinaire, 32) Then
            Drapeau = CodeBinaire - 32
          Else
            Drapeau = CodeBinaire
          End If
        Case "SPLINE"
          CodeBinaire = val(ValeurGroupe)
        'Vérifier si la spline est rationnelle
          mSpLine.Rational = BitActif(CodeBinaire, 4)

        'Exclure les splines périodiques
          If BitActif(CodeBinaire, 2) Then
            interessant = False
            Set mSpLine = Nothing
          Else
            mSpLine.ferme = BitActif(CodeBinaire, 1)
          End If
        End Select
      Case 10 ' point,circle,arc,insert,polyline,line,vertex
        If nomEntite = "LWPOLYLINE" Or nomEntite = "SPLINE" Then
          Sommets.Add val(ValeurGroupe) * echDXF - pBloc.x, 0#
        ElseIf nomEntite <> "POLYLINE" Then
          p0.x = val(ValeurGroupe) * echDXF - pBloc.x
        End If
      Case 20 ' point,circle,arc,insert,polyline,line,vertex
        If nomEntite = "LWPOLYLINE" Or nomEntite = "SPLINE" Then
          Sommets(Sommets.Count).y = val(ValeurGroupe) * echDXF - pBloc.y
        ElseIf nomEntite <> "POLYLINE" Then
          p0.y = val(ValeurGroupe) * echDXF - pBloc.y
        End If
      Case 11 ' line
        p1.x = val(ValeurGroupe) * echDXF - pBloc.x
      Case 21 ' line
        p1.y = val(ValeurGroupe) * echDXF - pBloc.y
      Case 40 'circle,arc,polyline,vertex
        If nomEntite = "ARC" Or nomEntite = "CIRCLE" Then Rayon = val(ValeurGroupe) * echDXF
        If nomEntite = "SPLINE" Then
          mSpLine.Noeuds(numNoeud) = val(ValeurGroupe)
          numNoeud = numNoeud + 1
        End If
      Case 41 ' insert,polyline,vertex
        If nomEntite = "INSERT" Then echx = val(ValeurGroupe)
        If nomEntite = "SPLINE" Then Sommets(Sommets.Count).arrondi = val(ValeurGroupe)
      Case 42 ' insert,lwpolyline,vertex,spline
        If nomEntite = "INSERT" Then
          echy = val(ValeurGroupe)
        ElseIf nomEntite = "VERTEX" Then
          arrondi(0) = val(ValeurGroupe)
        ElseIf nomEntite = "LWPOLYLINE" Then
          Sommets(Sommets.Count).arrondi = val(ValeurGroupe)
        ElseIf nomEntite = "SPLINE" Then
          mSpLine.tolNoeuds = val(ValeurGroupe)
        End If
      Case 44 'insert
        distcol = val(ValeurGroupe) * echDXF
      Case 45 'insert
        distrow = val(ValeurGroupe) * echDXF
      Case 50 'arc,insert (+ vertex : inutilisé dans la présente version de GIRATION)
        If nomEntite = "INSERT" Then
          rot = borne360(ValeurGroupe)
        Else
          Alpha(0) = borne360(ValeurGroupe)
        End If
      Case 51 ' ARC
        Alpha(1) = borne360(ValeurGroupe)
      Case 2  'INSERT
        nomBloc = ValeurGroupe
        If nomEntite = "INSERT" Then
''          If InStr(1, nomBloc, "GIR_", 1) = 1 Then nomBloc = nomBloc & "_$"
        ' vbTextCompare ne fait pas de différence Majuscule/Minuscule
          If InStr(1, nomBloc, "GIR_", vbTextCompare) = 1 And InStr(1, nomBloc, nomCourtVéhicule, vbTextCompare) = 5 Then
            VerifNomblocReserve nomBloc, Creation:=False
          End If
          If Not Graph.gphBlocs.Exist(nomBloc) Then interessant = False
        End If
      Case 67
        If val(ValeurGroupe) = 1 Then interessant = False     ' espace papier
      Case 71 'insert,polyline,spline
        If nomEntite = "INSERT" Then nbrow = val(ValeurGroupe)
        If nomEntite = "SPLINE" Then mSpLine.Ordre = val(ValeurGroupe) + 1
      Case 72 ' SPLINE
        nbNoeuds = val(ValeurGroupe)
      Case 73 ' SPLINE
        nbVertex = val(ValeurGroupe)
        Set Sommets = New Pts
      Case 75 'polyline lissée
        If nomEntite = "POLYLINE" Then
          pLine.TypeLissage = val(ValeurGroupe)
        End If
      Case 90 ' lwpolyline
        If nomEntite = "LWPOLYLINE" Then
          nbVertex = val(ValeurGroupe)
          'ReDim pv(nbVertex - 1)
          'ReDim arrondi(nbVertex - 1)
          'numVertex = -1
          Set Sommets = New Pts
        End If
      End Select
      
    End If  ' interessant
    
    'Même si ce n'est pas intéressant, on termine la boucle de lecture de l'entité
    lire_code -1
  Loop
  
  If echy <> echx Then interessant = False

  If interessant Then
    
    If nomPlan = "" And nomEntite <> "SEQEND" Then
      ErreurFDP IDm_EntiteSansPlan
    End If
    
    ' Forçage de la couleur : lors d'un futur Export, les nouveaux objets seront dans le même plan GIR_<nomvehicule>_xxx
    ' Comme ce plan aura la couleur des options d'affichage, on perdrait les couleurs d'origine du DXF
    If InStr(5, nomPlan, nomCourtVéhicule, vbTextCompare) = 5 And couleur = 256 Then
      couleur = gtPlans(nomPlan).couleur
    End If
    
    If objetBloc Is Nothing Then
      With Graph
        Select Case nomEntite
        Case "POINT"
          .gphPoints.Add nomPlan, couleur, p0
        Case "LINE"
          .gphLines.Add nomPlan, couleur, typelign, p0, p1
        Case "LWPOLYLINE"
          .gphPolylines.AddPline pLine
        Case "SPLINE"
          .gphSplines.Add mSpLine
        Case "POLYLINE"
          .gphPolylines.AddPline pLine
        Case "CIRCLE"
          .gphCercles.Add nomPlan, couleur, typelign, p0, Rayon
        Case "ARC"
          .gphArcs.Add nomPlan, couleur, typelign, p0, Rayon, Alpha(0), Alpha(1)
        Case "INSERT"
          .gphInserts.Add nomBloc, nomPlan, couleur, typelign, p0, echx, echy, rot, nbcol, nbrow, distcol, distrow
        Case "SEQEND"
          Set mPolylines = .gphPolylines
        End Select
      End With  ' objet Graph (fCourante.GraphFDP)
     
    Else    ' Entité membre d'un bloc
      With objetBloc
        Select Case nomEntite
        Case "POINT"
          .blkPoints.Add nomPlan, couleur, p0
        Case "LINE"
          .blkLines.Add nomPlan, couleur, typelign, p0, p1
        Case "LWPOLYLINE"
          .blkPolylines.AddPline pLine
        Case "SPLINE"
          .blkSplines.Add mSpLine
        Case "POLYLINE"
          .blkPolylines.AddPline pLine
        Case "CIRCLE"
          .blkCercles.Add nomPlan, couleur, typelign, p0, Rayon
        Case "ARC"
          .blkArcs.Add nomPlan, couleur, typelign, p0, Rayon, Alpha(0), Alpha(1)
        Case "INSERT"
          .blkInserts.Add nomBloc, nomPlan, couleur, typelign, p0, echx, echy, rot, nbcol, nbrow, distcol, distrow
        Case "SEQEND"
          Set mPolylines = .blkPolylines
        End Select
        
      End With  ' bloc
    
    End If  ' objetBloc Is Nothing
  
    If Not pLine Is Nothing Then
      Select Case nomEntite
      Case "POLYLINE", "LWPOLYLINE"
        With pLine
          .couleur = couleur
          .typelign = typelign
          .nomPlan = nomPlan
          If nomEntite = "LWPOLYLINE" Then
            For numVertex = 1 To nbVertex
              If numVertex <> 1 Then
                If Sommets(numVertex - 1).arrondi <> 0 Then CalArc pLine, Sommets(numVertex), nomPlan, couleur, typelign
              End If
              .Pts.AddPt Sommets(numVertex)
            Next
            
            If .ferme Then
              If Sommets(nbVertex).arrondi <> 0 Then
                CalArc pLine, Sommets(1), nomPlan, couleur, typelign
              End If
            End If
            'Ajout AV (14/01/04 : v3.2.33) : pouvait faire planter ultérieurement un insert de bloc avec attribut(entité SEQEND à ignorer alors )
            Set pLine = Nothing
          End If
        End With
        
      Case "VERTEX"
        If flagArrondi Then   ' création de l'Arc avec le pt précédent
          CalArc pLine, p0, nomPlan, couleur, typelign
        End If
        pLine.Pts.Add p0.x, p0.y, arrondi(0), Drapeau
        flagArrondi = arrondi(0) <> 0
        'Si on veut retracer la polyligne splinée, il faut également (re)mettre en commentaire : 'Coul = QBColor(tCouleur(Abs(Coul)))' dans CERCLE.Dessiner
'          If (Drapeau And 16) <> 16 Then .Cercles.Add nomPlan, QBColor(12), typelign, p0, 1
      Case "SEQEND"
        With mPolylines
          If pLine.Pts.Count < 2 Then
            ErreurFDP IDm_UnSeulPointPline
          ' on fait comme AutoCAD : on arrête le chargement (suite ignorée)
            .Remove .Count
          End If
          
          If pLine.ferme And flagArrondi Then CalArc pLine, pLine.Pts(1), nomPlan, couleur, typelign
        End With  ' mPolylines
        
        flagArrondi = False
        Set pLine = Nothing
      End Select
    End If

    If Not mSpLine Is Nothing Then
      With mSpLine
        .couleur = couleur
        .typelign = typelign
        .nomPlan = nomPlan
        For Each pLoop In Sommets
          .PtsControle.AddPt pLoop
        Next
        .Etablir nbPoints:=50
      End With
      Set mSpLine = Nothing
    End If
  
  ElseIf nomEntite = "POLYLINE" Or nomEntite = "LWPOLYLINE" Then
    Set pLine = Nothing
  End If   ' interessant
 
Exit Sub

GestErr:

  ErrLocal Procédure:="LecDXF.lire_entite"
  Err.Raise Err.Number
End Sub

'******************************************************
' Recherche si la polyligne est exclue par GIRATION
'******************************************************
Private Function CodeExcluPoly(ByVal CodeBinaire As Integer) As Boolean

'  If BitActif(CodeBinaire, 2) Then
' Bit 1 : Courbe lissée
'    CodeExcluPoly = True
'
'  ElseIf BitActif(CodeBinaire, 4) Then
'' Bit 2 : Spline
'    CodeExcluPoly = True
    
  If BitActif(CodeBinaire, 16) Then
' Bit 4 : Maillage 3D
    CodeExcluPoly = True
    
  ElseIf BitActif(CodeBinaire, 32) Then
' Bit 5 : Maille polygone fermée
    CodeExcluPoly = True
    
' Bit 6 : Maille polyface
  ElseIf BitActif(CodeBinaire, 64) Then
    CodeExcluPoly = True
  End If
  
End Function

'******************************************************
' Recherche si le sommet est exclu par GIRATION
'******************************************************
Private Function CodeExcluSommet(ByVal CodeBinaire As Integer) As Boolean

' Bit 0 - Drapeau=1  - Sommet créé par lissage (polylige lissée, ou polyline splinée avec SPLINESEGS <0)
' Bit 2 - Drapeau=4  - Inutilisé
' Bit 3 - Drapeau=8  - Sommet spline créé par lissage (polyligne splinée uniquement)
' Bit 4 - Drapeau=16 - Point de controle de spline  (polyligne splinée uniquement)
' Bit 5 - Drapeau=32 - Sommet de polyligne 3D
  
  'If BitActif(CodeBinaire, 2) Then
' Bit 1 : Tangence à Courbe lissée !! Giration ne sait pas interpréter le code 50
   ' CodeExcluSommet = True
    
  'If BitActif(CodeBinaire, 16) Then
' Bit 4 : Point de controle - le sommet ne doit pas être dessiné,
' mais être conservé en mémoire pour réécriture : Voir Calarc et POLYLINE.PtsUtiles
  '  CodeExcluSommet = True
    
  If BitActif(CodeBinaire, 64) Then   ' ne doit jamais arriver vu que les polylignes correspondantes ont été exclues
' Bit 6 : Maillage 3D
    CodeExcluSommet = True
    
' Bit 7 : Maille polyface
  ElseIf BitActif(CodeBinaire, 128) Then   ' ne doit jamais arriver vu que les polylignes correspondantes ont été exclues
    CodeExcluSommet = True
  End If
  
End Function



'******************************************************
' Détermine si le bit vaut 0 ou 1 dans CodeBinaire
'******************************************************
Public Function BitActif(ByVal CodeBinaire As Integer, ByVal NumBit As Integer) As Boolean
Dim ValeurBit As Integer

  'ValeurBit = 2 ^ PositionBit  : utile si la position du bit est fournie plutôt que la valeur recherchée
  ValeurBit = NumBit
  
  BitActif = ((CodeBinaire And ValeurBit) = ValeurBit)

End Function

'******************************************************************************
' Décryptage du type de données selon son code (fichier DXF binaire)
'******************************************************************************
Private Sub DecrypSelonCode()
Dim typdonn As TypeDonnBinaire
Dim Msg As String
Dim i%

On Error GoTo GestErr

If Binaire Then

  Select Case CODE
  Case 0 To 9, 998 To 1009
          typdonn = typeChaine
          If CODE = 1004 And Binaire Then typdonn = typeEtendueBinaire 'données d'entités étendues binaires
  Case 10 To 59
          If CODE = 41 Or CODE = 42 Or CODE > 49 Then
            typdonn = typeDouble
          Else
            typdonn = typeCoord
          End If
  Case 60 To 79, 170 To 178, 1060 To 1079   ' 175-->Repoussé à 178 pour les textes (bug doc DXF) - AV - 14/01/99
          typdonn = typeEntier
          If CODE = 1071 Then typdonn = typeLong 'données d'entités étendues de type entier long
  Case 90 To 99   ' AutoCAD 14
          typdonn = typeLong
  Case 100, 102, 105, 300 To 369      ' AutoCAD 14
          typdonn = typeChaine
          If Binaire And CODE > 309 And CODE < 320 Then
            typdonn = typeEtendueBinaire
          End If
  Case 110, 111, 112, 120, 121, 122, 130, 131, 132  ' non documenté dans ref DXF 2000 mais trouvé dans un fichier du CERTU (26/09/03)
          typdonn = typeCoord
  Case 390 To 399, 410 To 419          ' AutoCAD 2000
          typdonn = typeChaine
  Case 270 To 275, 280 To 289     ' AutoCAD 14 - 270 à 275 non documenté mais présent dans DIMSYTLE
          typdonn = typeEntier
  Case 290 To 299       ' AutoCAD 2000 (codes 290-299 rajoutés par AV le 26/09/03 : v 3.2.32 - Doc DXF incomplète et en + on n'indique pas dans l'annexe sur les fichiers binaires qu'un booléen est stocké sur un octet")
          typdonn = typeBooléen
  Case 370 To 389, 400 To 409        ' AutoCAD 2000
          typdonn = typeEntier
  Case 80 To 139, 148 To 169, 179 To 209, 240 To 997
          typdonn = False
  Case 140 To 147, 210 To 239, 1010 To 1059
          typdonn = typeDouble
  End Select
  
  Select Case typdonn
    Case typeBooléen
      lire_ligne "Booléen"
    Case typeChaine
      lire_ligne "chaine"
    Case typeEntier
      lire_ligne "entier"
    Case typeLong
      lire_ligne "long"
    Case typeDouble
      lire_ligne "double"
    Case typeCoord  ' coordonnées de pt auxquelles on applique un facteur d'échelle(/ unités AutoCAD)
      lire_ligne "double"
    Case typeEtendueBinaire   ' Données binaires dans un fichier DXF binaire(code 1004)
      lire_ligne "octet"  ' le 1er octet contient le nombre d'octets à lire
      For i = 1 To val(CodeGroupe): lire_ligne "octet": Next
    Case Else
      lire_ligne "chaine"
      Msg = ID_Code & " " & CStr(CODE) & " " & IDm_Incorrect
      ErreurFDP Msg, NonBloquant:=True
    End Select
  End If
  
Exit Sub
GestErr:
  ErrLocal Procédure:="LecDXF.DecypSelonCode"
End Sub

'******************************************************************************
' Lecture du code suivant
'******************************************************************************
Private Sub lire_code(ByVal controle As Integer, Optional ByVal CodeFacultatif As Integer = -2)
' controle représente la valeur du code attendu, -1 indique que le code peut être quelconque
Dim octet As Byte
Static CoordXTraité As Boolean

On Error GoTo GestErr

    If Not Binaire And (CODE = 10 Or CODE = 11 Or CODE = 12 Or CODE = 13) And Not CoordXTraité Then
      CodeGroupe = CodeGroupe2
      ValeurGroupe = ValeurGroupe2
      CoordXTraité = True
    ElseIf CoordXTraité And (CODE = 20 Or CODE = 21 Or CODE = 22) Then
      CodeGroupe = CodeGroupe3  ' CodeGroupe3 contient le Z ou le code groupe qui suit si le point est un point 2D
      ValeurGroupe = ValeurGroupe3
    Else
      CoordXTraité = False
    If dwgVersion = 12 Then
      lire_ligne "octet"
    Else
      lire_ligne "CodeGroupe" ' Code sur 2 octets à partir de v13 (non documenté)
    End If
  End If

  CODE = val(CodeGroupe)
  DecrypSelonCode   '(==Ajout)
  
  Select Case CODE
  'Ajout AV : 06/11/2000 - Code 102
  Case 102
    '  Données attachées à une entité : incluses entre 2 accolades {}
'          If ValeurGroupe <> "}" Then
    ' Correction AV ! 22/11/2001 : L'objet XRECORD peut utiliser le code 102 pour une donnée applicative simple (donc pas d'accolades)
  '(==Suppression)  DecrypSelonCode
    If Left(ValeurGroupe, 1) = "{" Then 'Ex. "{ACAD_XDICTIONARY"  "{ACAD_REACTORS"
        'lire toutes les données jusqu'à retrouver l'accolade fermante
      Do
        lire_code -1
  '(==Suppression)      If CODE <> 102 Then DecrypSelonCode
      Loop Until CODE = 102
      ' Rappel récursif de lire_code pour traiter le cas qui nous intéressait au départ
      lire_code controle, CodeFacultatif
      ' Arrêt : La suite du traitement a été effectuée lors de l'appel récursif
    ElseIf Left(ValeurGroupe, 1) = "}" Then 'Ex. "{ACAD_XDICTIONARY"  "{ACAD_REACTORS"
      controle = -1
    End If
    
  Case 5, 105, 330, 100, 340
        ' code 340 ajouté (AV : 26/09/03) : présence inattendue d'un code 340 dans 1 fichier particulier
    lire_code controle, CodeFacultatif
    
  End Select

  If (controle <> -1) Then
    If CodeFacultatif = -2 Then
      If CODE <> controle Then
        ErreurFDP IDm_DXFIncorrect & " (" & ID_Attendu & " : " & CStr(controle) & ")"
      End If
    ElseIf CODE <> controle And CODE <> CodeFacultatif Then
      ErreurFDP IDm_DXFIncorrect & " (" & ID_Attendu & " : " & CStr(controle) & ")"
    End If
  End If

Exit Sub
GestErr:
  
  If Err <> 100 Then
    ErrLocal Procédure:="LecDXF.lire_code"
  Else
    Err.Raise 100
  End If

End Sub

'******************************************************************************
' Lecture dune chaine
'******************************************************************************
Private Sub lire_chaine(ByVal controle As String, ParamArray arg())

Dim chaineOk As Boolean, i As Integer
Dim mess As String

  mess = controle
  chaine = ValeurGroupe

  If controle = chaine Then
    chaineOk = True
  ElseIf controle = "" Then
    chaineOk = True
  Else
    For i = LBound(arg) To UBound(arg)
      mess = mess & " ou " & arg(i)
      If chaine = arg(i) Then chaineOk = True: Exit For
    Next i
  End If

  If Not chaineOk Then
    mess = remplaceOuParVirgule(mess)
    ErreurFDP "'" & chaine & "' " & IDm_Incorrect & " (" & ID_Attendu & " : " & mess & ")"
  End If

End Sub

'******************************************************************************
' Remplacement de tous les "ou" de l'énumération par une virgule, à l'exception du dernier
'******************************************************************************
Private Function remplaceOuParVirgule(s As String) As String
Dim pos%

remplaceOuParVirgule = s

pos = InStr(s, " ou ")
If pos <> 0 Then
  If InStr(pos + 4, s, " ou ") <> 0 Then  ' il y a encore un "ou" plus loin dans la chaine
    remplaceOuParVirgule = Left(s, pos - 1) & ", " & remplaceOuParVirgule(Mid(s, pos + 4))
  End If
End If

End Function

'******************************************************************************
' Lire une ligne du fichier
'******************************************************************************
Private Sub lire_ligne(ByVal param As String)
Dim entier As Integer, l As Long, dble As Double, carac As String * 1, s As String, octet As Byte
Dim PosEnCours As Long

'    If PourCentEnCours = 0 Then frmDéfileDXF.lblpourcent = "0%"

  If Binaire Then
    Select Case param
    Case "Booléen"
      Get #numFich, , octet
      
    Case "CodeGroupe"   ' v13 et suivantes (non documenté) : code groupe sur 2 octets
      Get #numFich, , entier
      CodeGroupe = CStr(entier)
    Case "octet"
      Get #numFich, , octet
      If dwgVersion = 12 And octet = 255 Then       ' données d'entités étendues  ( v12 uniquement, vu commentaire ci-dessus)
        lire_ligne "entier"
        CodeGroupe = ValeurGroupe
      Else
        CodeGroupe = CStr(octet)
      End If
    Case "entier"
      Get #numFich, , entier
      ValeurGroupe = CStr(entier)
    Case "long"
      Get #numFich, , l
      ValeurGroupe = CStr(l)
    Case "double"
      Get #numFich, , dble
      ValeurGroupe = substPtDecimalRegional(CStr(dble), Regional:=False)
    Case "chaine"
      Get #numFich, , carac
      While carac <> Chr(0)
        s = s + carac
        Get #numFich, , carac
      Wend
      ValeurGroupe = s
    End Select   ' Case param
    If EOF(numFich) Then
      ErreurFDP IDm_FinPrematuree
    Else
      PosEnCours = Loc(numFich)
    End If

  Else    ' ASCII
''''''''''''''''''''''      Line Input #numFich, codegroupe
    Input #numFich, CodeGroupe
    num_ligne = num_ligne + 1
    If Not IsEntier(CodeGroupe) Then
        ErreurFDP IDm_DXFIncorrect & " (" & ID_NombreEntier & " " & ID_Attendu & ")"
    End If
    If CodeGroupe = "10" Or CodeGroupe = "11" Or CodeGroupe = "12" Or CodeGroupe = "13" Then
      ' coordonnée X lue : on lit le Y et le Z qui en principe suivent juste après(si c'est un point 2D la lecture du 3ème Groupe sera gérée par lire_code)
        Input #numFich, ValeurGroupe, CodeGroupe2, ValeurGroupe2, CodeGroupe3, ValeurGroupe3
        num_ligne = num_ligne + 5
    Else
      Line Input #numFich, ValeurGroupe
      ValeurGroupe = DécodeUTF8(ValeurGroupe)
      num_ligne = num_ligne + 1
    End If
    PosEnCours = Loc(numFich) * 128
  End If
    
        ' Affichage du pourcentage suivant
  While PosEnCours >= (PourCentEnCours + 1) * UnPourCent ' Then
    frmDéfileDXF.lblPourCent = PourCentEnCours + 1 & "%"
    PourCentEnCours = PourCentEnCours + 1
    DoEvents
  Wend

  If frmDéfileDXF.Annul Then Err.Raise 102
    
End Sub

Private Function DécodeUTF8(ByVal chaine As String) As String
Dim Retour As String
Dim pos As Integer
Dim carac As String
Dim valAscii As Integer

  pos = InStr(chaine, "Ã")
'  Exit Function
  If pos = 0 Then
    DécodeUTF8 = chaine
  Else
    Do
      carac = Mid(chaine, pos + 1)
      valAscii = Asc(carac)
      Select Case valAscii
'      Case 160
'        carac = "à"
'      Case 167
'        carac = "ç"
'      Case 168
'        carac = "è"
'      Case 169
'        carac = "é"
'      Case 185
'        carac = "ù"
      Case 128 To 156
      ' caractères accentués majuscules
        carac = Chr(valAscii + 64)
      Case 160 To 188
      ' caractères accentués minuscules
        carac = Chr(valAscii + 64)
      Case Else
      ' caractère sans équivalent en ascii étendu : mettre un underscore
        carac = "_"
      End Select
      
      DécodeUTF8 = DécodeUTF8 & Left(chaine, pos - 1) & carac
      chaine = Mid(chaine, pos + 2)
      pos = InStr(chaine, Chr(195))
    Loop Until pos = 0
    DécodeUTF8 = DécodeUTF8 & chaine
  End If
  
End Function

Public Sub minimax(ParamArray coord())
Dim i As Integer, nbcoord As Integer

nbcoord = (UBound(coord) - LBound(coord) + 1)
For i = 1 To nbcoord Step 2
  With Wpmin
    If coord(i - 1) < .x Then .x = coord(i - 1)
    If coord(i) < .y Then .y = coord(i)
  End With
  With Wpmax
    If coord(i - 1) > .x Then .x = coord(i - 1)
    If coord(i) > .y Then .y = coord(i)
  End With
Next
  
End Sub

Public Sub calminmax(Optional ByVal objBloc As BLOC, Optional ByVal pInsert As PT)
Dim Cpt As Integer, cpt2 As Integer, p As POINT
Dim xIns As Double, yIns As Double
Dim Angle As Double
Dim i As Integer
Dim eqv180 As Double

Dim objGraph As GRAPHIQUE
Dim NewPInsert As PT

Dim colPoints As Points
Dim colLines As Lines
Dim colPLines As Polylines
Dim colArcs As Arcs
Dim colCercles As Cercles
Dim colInserts As Inserts

Dim objPoint As POINT
Dim objLine As LINE
Dim objPline As POLYLINE
Dim objArc As ARC
Dim objCercle As CERCLE
Dim objInsert As INSERT

On Error GoTo GestErr

  eqv180 = eqvPI(OptGen.unite)
  Set objGraph = Graph
  
  If objBloc Is Nothing Then
    With objGraph
      Set colPoints = .gphPoints
      Set colLines = .gphLines
      Set colPLines = .gphPolylines
      Set colArcs = .gphArcs
      Set colCercles = .gphCercles
      Set colInserts = .gphInserts
    End With
    
  Else
    With objBloc
      Set colPoints = .blkPoints
      Set colLines = .blkLines
      Set colPLines = .blkPolylines
      Set colArcs = .blkArcs
      Set colCercles = .blkCercles
      Set colInserts = .blkInserts
    
      'Point d'insertion
      xIns = pInsert.x
      yIns = pInsert.y
    End With
  End If
  
  
  For Each objPoint In colPoints
    With objPoint
      minimax .p.x + xIns, .p.y + yIns
    End With
  Next
  
  For Each objLine In colLines
    With objLine
      minimax .Pts(1).x + xIns, .Pts(1).y + yIns, .Pts(2).x + xIns, .Pts(2).y + yIns
    End With
  Next
  
  For Each objPline In colPLines
    With objPline
      For cpt2 = 1 To .Pts.Count
        minimax .Pts(cpt2).x + xIns, .Pts(cpt2).y + yIns
      Next
    End With
  Next
  
  For Each objCercle In colCercles
    With objCercle
      minimax .p.x - .Rayon + xIns, .p.y - .Rayon + yIns, .p.x + .Rayon + xIns, .p.y + .Rayon + yIns
    End With
  Next
  
  For Each objArc In colArcs
    With objArc
      If .angdep < .angfin Then
        Angle = .angdep
      Else
        Angle = .angdep - eqv180 * 2
      End If
      While Angle < .angfin
        minimax .p.x + xIns + .Rayon * Cos(Angle * pi / (eqv180)), .p.y + yIns + .Rayon * Sin(Angle * pi / (eqv180))
        Angle = Angle + eqv180 / 6
      Wend
      minimax .p.x + xIns + .Rayon * Cos(.angfin * pi / (eqv180)), .p.y + yIns + .Rayon * Sin(.angfin * pi / (eqv180))
    End With
  Next
  
  For Each objInsert In colInserts
    With objInsert
      Set NewPInsert = New PT
      NewPInsert.x = xIns + .p.x
      NewPInsert.y = yIns + .p.y
      calminmax objBloc:=objGraph.gphBlocs(.name), pInsert:=NewPInsert
      Set NewPInsert = Nothing
    End With
  Next

Exit Sub
GestErr:
  Debug.Print
End Sub

Public Sub reinFDP()
Dim objPlan As PLAN

  For Each objPlan In gtPlans
     If InStr(objPlan.name, "$$$GIR_") = 0 Then
      gtPlans.Remove objPlan.name
    End If
 Next
  
 If Not Graph Is Nothing Then
    With Graph
      With .gphPoints
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
      With .gphLines
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
      With .gphPolylines
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
      With .gphCercles
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
      With .gphArcs
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
      With .gphInserts
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
      With .gphBlocs
        Do While .Count <> 0
          .Remove .Count
        Loop
      End With
    End With
  End If
    
End Sub

Private Sub CalArc(ByVal pLine As POLYLINE, ByVal p2 As PT, ByVal nomPlan As String, ByVal couleur As Long, ByVal typelign As String)
Dim p1 As PT, centre As New PT, Rayon As Double, angdeb As Double, angfin As Double
'Calcul du centre d'un arc de polyligne défini par ses extrémités et par son arrondi
' L'arrondi est la tangente du quart de l'arc décrit.
' Le signe de l'arrondi est positif si l'arc est décrit dans le sens trigonométrique
' négatif dans le cas contraire
' Calcule également le rayon de l'arc ainsi que l'angle de début et de fin au sens de la méthode graphique CIRCLE de VB

Dim d As Double, x0 As Double, y0 As Double, xc As Double, yc As Double, Alpha As Double, alphap As Double, Beta As Double, Gamma As Double
Dim arrondi As Double
Dim mPts As Pts

    'PtsUtiles : Pts de la polyligne sauf les points de controle en cas de polyligne Splinée
    Set mPts = pLine.PtsUtiles
    
    ' p1 : pt précédent
     Set p1 = mPts(mPts.Count)
     arrondi = p1.arrondi
               'calcul de la distance entre les 2 sommets consécutifs
     d = Distance(p1.x, p2.x, p1.y, p2.y)
               ' rayon
     Rayon = Abs((d / 2) * (1 + Carre(arrondi)) / (2 * arrondi))
             ' arc alpha
     Alpha = 4 * Atn(Abs(arrondi))
     alphap = Alpha
     If alphap > pi Then alphap = 2 * pi - alphap  ' => 0 < alphap < pi
     If arrondi < 0 Then Alpha = -Alpha                ' => -2pi < alpha < 2pi
      Beta = pi / 2 - alphap / 2
     If Alpha > pi Or (Alpha > -pi And Alpha < 0) Then Beta = -Beta
             ' centre du cercle
     x0 = d / 2
     y0 = Rayon * Sin(Beta)
     Gamma = CalcAngle(p1, p2)
     centre.x = x0 * Cos(Gamma) - y0 * Sin(Gamma) + p1.x
     centre.y = x0 * Sin(Gamma) + y0 * Cos(Gamma) + p1.y
       ' Circle dessine tjs dans le sens trigo
     
     If arrondi > 0 Then ' sens trigo
       angdeb = CalcAngle(centre, p1)
       angfin = CalcAngle(centre, p2)
     Else                        ' sens horaire
       angdeb = CalcAngle(centre, p2)
       angfin = CalcAngle(centre, p1)
     End If
     ' La méthode Circle utilise le signe négatif d'un angle pour tracer un rayon depuis le centre
     ' et ignore le signe
     If angdeb < 0 Then angdeb = angdeb + 2 * pi
     If angfin < 0 Then angfin = angfin + 2 * pi
     angdeb = angdeb * 180# / pi
     angfin = angfin * 180# / pi

     pLine.Arcs.Add nomPlan, couleur, typelign, centre, Rayon, angdeb, angfin

End Sub

Private Function borne360(v As String) As Double
' Pour VB, un angle doit être dans l'intervalle [0,360[
' le signe - ayant une signification particulière pour la méthode graphique CIRCLE

  borne360 = val(v)
  If val(borne360) < 0 Then borne360 = borne360 + 360
  
' Dans les fichiers DXF, les angles sont toujours en degrés : il y a lieu de les transformer si l'utilisateur de GIRATION travaille en grades
  If OptGen.unite = GRADE Then borne360 = borne360 / 0.9

End Function

Private Function IsEntier(s As String) As Boolean
  
On Error GoTo GestErr

    IsEntier = (CStr(CInt(s)) = s)
    Exit Function
    
GestErr:
    IsEntier = False
    Exit Function
    
End Function
