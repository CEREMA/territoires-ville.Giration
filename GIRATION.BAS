Attribute VB_Name = "GirationMain"
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97
'     Mise à jour pour la version anglaise et la gestion des imprimantes sous Windows NT : Février 2000
'

'   Réalisation : André VIGNAUD

'   Module standard : Giration   -   GIRATION.BAS
'   Module de lancement de GIRATION
'   Fonctions du module
'       Main  : Ouverture de la feuille MDIGiration
'                     Lecture de la bibli de véhicules (lireVeh)
'                     Initialisations
'       GriserMenus : Activation ou non des  options de menu selon le contexte
'       Gestions de fichiers : Ouvrir,Enregistrer,Import,Export,
'**************************************************************************************
Option Explicit

'  #Const GIRATIONDEMO = True

' Déclaration des variables globales
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'Public gtProtectRetirée As Integer

Public gtBibliVerrouillée As Boolean

Public gtFso As Scripting.FileSystemObject
Public gtFeuillesTraject As New Collection

Public numFichLog As Integer, FichierJournal As Boolean

Public VersionDemo As Boolean, VersionDemoNom As String * 20
Public gtVersionDéveloppeur As Boolean
Public gtAppliDistante As Boolean

Public ErreurGenerale As Boolean

Public gtDistance As Boolean
Public gtAngleDyn As Boolean
'Public gtAngleOri As Boolean

Public gtBoutons(3) As String
Public gtOutil As TOOLEnum
Public gtOutilZoom As TOOLZOOMEnum
Public gtCoordActif As Boolean

Public Enum TOOLEnum
  TOOL_AUCUN = -1
  TOOL_COORDDYN = 0
  TOOL_XY = 1
  TOOL_DIST = 2
  TOOL_ANGLEDYN = 3
End Enum

Public Enum TOOLZOOMEnum
  TOOL_SANSZOOM = -1
  TOOL_ZOOM = 0
  TOOL_PAN = 1
  TOOL_ZAVANT = 2
  TOOL_ZARRIERE = 3
  TOOL_ZTOUT = 4
End Enum

'Public Const TOOL_COORDDYN = 0
'Public Const TOOL_XY = 1
'Public Const TOOL_DIST = 2
'Public Const TOOL_ANGLEDYN = 3
'Public Const TOOL_ZOOM = 4
'Public Const TOOL_PAN = 5
'Public Const TOOL_ZAVANT = 6
'Public Const TOOL_ZARRIERE = 7
'Public Const TOOL_ZTOUT = 8

Public Const ErrVéhicule = 1100

Public GirInitOk As Boolean    ' Drapeau lors de l'initialisation de GIRATION

'Public flagVirgule As Boolean
Public alertVirgule As Boolean
Public gtPtDecimal As Integer
'Public gtptdecimal As String

Public grisPale As Long, blanc As Long
' Indique que le module d'impression frmImprim est chargé
Public flagImpr As Boolean
' Indique que la procédure Export est en cours
Public flagExport As Boolean

'Public Const GirationVersion = "3.3"
Public GirationVersion As String

' Feuilles fille MDI frmTraject
'---------------------------------------

'Nombre d'instances créées de la fenêtre frmTraject.
Public NbreFilles As Integer, NumFille As Integer

'Fenêtre trajectoire courante
Public fCourante As frmTraject

' Dessin de l'épure
'-------------------------
'Objet recevant l'épure en cours
Public ObjetDessin As Object
Public sensTracteur As Boolean
' Appel des fonctions de dessin pour calculer l'échelle, sans dessiner
Public pourEchelle As Boolean

'Constantes des rubriques d'aide
'--------------------------------------------
Public Const HELPNAME = "Giration.chm"
Public Const SOMMAIRE = 0

Public Const CARACVEHICUL = 102       ' frmVéhicule           : 2.3.2

Public Const BIBVEHICUL = 204       ' frmBibVehicule + menu : 2.3.1.1
Public Const NEWVEHICUL = 204       ' frmNewVehicule
Public Const IMPORTVEHICUL = 206    ' frmImportBib + menu   : 2.3.1.2
Public Const EXPORTVEHICUL = 207    ' frmExportBib + menu   : 2.3.1.3
Public Const DEFVEHICUL = 209     ' frmVéhicule (depuis fenêtre Trajectoire)

'Public Const GENERALITESTRAJ = 214  ' Présentation générale : 2.4.1
Public Const GENERALITESTRAJ = 221  ' Présentation générale : 2.4.3.2
Public Const SAISORI = 220          ' frmTraject : BasculeGrilleOrigine : 2.4.3.1

Public Const NOUVTRAJ = 216         ' frmTrajpar    + menu  : 2.4.2.1
Public Const OUVRIRTRAJ = 217       ' Ouvrir trajectoire    : 2.4.2.2
Public Const ENREGTRAJ = 218        ' Enregistrer trajectoire: 2.4.2.3
Public Const EXPORTCARAC = 224      ' Exporter les caractéristiques: 2.4.3.5

Public Const OPTAFFICHAGE = 227     ' frmOptAff    + menu   : 2.4.4.1
Public Const CADRAGE = 228          ' frmCadrage    + menu : 2.4.4.2
Public Const IMPRIMER = 229         ' frmImprim    + menu   : 2.4.4.3
Public Const PASAPAS = 230          ' frmPas        + menu  : 2.4.4.4
Public Const LARGEUR = 232          ' frmLargeur    + menu  : 2.4.4.6
Public Const MESURES = 233          ' Outils de mesure      : 2.4.4.7

Public Const IMPORTFDP = 236        ' frmImport     + menu  : 2.4.5.1
Public Const EXPORTEPURE = 237      ' Exporter l'épure        : 2.4.5.2
Public Const ENREGFDP = 238         ' Enregistrer fond de plan: 2.4.5.3
Public Const OPTGENERAL = 301       ' frmOptGen    + menu   : 3.1

'Constantes caduques en v3.3
'Public Const CONFIGIMPR = 233      ' Configuration de l'imprimante
'Les 3suivants n'ont plus de fenêtres spécifiques : appel de l'aide sur la grille de saisie
Public Const INSERSUPPRDEPL = 221  ' Insertion/Suppression de déplacement
Public Const DECOMPVIRAGE = 221    ' Décomposition de virage
Public Const ANGLETOTAL = 221      ' Affichage totalisation des angles


'Constantes des items de menus
'--------------------------------------------
' ---- Menu Fichier
Public Const MNUNOUVTRAJ = 1
Public Const MNUOUVRIRTRAJ = 2
Public Const MNUENREGTRAJ = 4
Public Const MNUENREGTRAJSOUS = 5
Public Const MNUEXPORTEPURE = 7
Public Const MNUEXPORTCARAC = 8
Public Const MNUBIBVEH = 10
Public Const MNUEXPORTBIB = 11
Public Const MNUIMPORTBIB = 12
Public Const MNUIMPORTFDP = 14
Public Const MNUENREGFDP = 15
Public Const MNUCONFIGIMPR = 17
Public Const MNUIMPRIMER = 18
' ---- Menu Outils
Public Const MNUINSERDEPL = 1
Public Const MNUSUPPRDEPL = 2
Public Const MNUDECOMPVIRAGE = 3
Public Const MNUPASAPAS = 5
Public Const MNULARGEUR = 6
Public Const MNUCADRAGE = 8
Public Const MNURAFRAICHIR = 9
Public Const MNUOPTGENERAL = 11
Public Const MNUOPTAFFICHAGE = 12
Public Const MNUMESURES = 14
Public Const MNUANGLETOTAL = 15

Public Const GIRATIONVEHCOURT = "Giration.veh"

' Mise en grisé d'un label
'Public Const grise = &H808080

' Noms des fichiers GIRATION
'-----------------------------------------
' Référence: Cahier des charges p3
Public GirationIni As String       ' GIRATION.INI
Public GirationVeh As String    '  GIRATION.VEH : bibiothèque de véhicules
Private DefautRep(1) As String

'Noms des plans réservés GIRATION
'---------------------------------------------------
Public GirPlanReserve(7) As String ' $$$GIR_n
Public radicExport(6) As String
Public PlanExport(6) As String ' Plan réservé  à l'export des carrosseries et des roues
Public BlocExport(2) As String  ' Bloc réservé  à l'export des carrosseries
Public ExportPlans As New Plans

' Derniers fichiers Trajectoire (.GIR°) utilisés
'-----------------------------------------------------------
Public MRUFichiers() As String, nbfichMenu As Integer

' Autres constantes (Utilisées dans frmTraject et frmVéhicule)
Public Const TRACTEUR = 0, REMORQUE1 = 1, REMORQUE2 = 2

' Unités d'angle
'--------------------
Public Const DEGRE = 0
Public Const GRADE = 1

' Booléens pour conversion radian-->unité utilisateur et réciproquement (fonction angConv dans GIRSTAND.BAS)
Public Const radian = True
Public Const unite = False

Public libUnite(1) As String, eqvPI(1) As Integer

' Véhicules
'--------------

Public VerrouVéhicule As Boolean  ' Verrouillage autorisé pour le CERTU seulement

' ColVéhicules est une collection d'objets de la classe VEHICULE
Public ColVéhicules As New Collection

' Format d'enregistrement d'un Véhicule
Public Type StructVéhicule
  Nom As String * 20
  Protégé As Boolean
  nbRemorque As Integer
  PAFav(2) As Single
  e(2) As Single
  PAFar(2) As Single
  l(2) As Single
  LEav As Single
  LEar(2) As Single
  Dpat(1) As Single
  alphaVehMax As Single
  aVehMax As Single
  RVehMin As Single
End Type

' version de bibliothèque (susceptible de changer moins souvent que la version du programme)
Public vehicVersion As StructVéhicule

'Options générales
'-------------------------
Public Type typOptGen
  unite As Integer
  Repert(2) As String
  VitMax As Single
  VitDef As Single
  Epaisseur As Integer
End Type
Public OptGen As typOptGen

' Mémorisation d'un import DXF pour l'import suivant dans la même session
Public gtRepertFDP As String

Public Type caracdess
  ADessiner As Boolean
  couleur As Long
End Type

' Options d'affichage
Public Type typOptAff
  Carross(2) As caracdess
  RouAvt As caracdess
  RouArr(2) As caracdess
  FDP As caracdess
  Graduations As caracdess
  pas(2) As Single
  sens As Boolean
  position As Boolean
End Type

Public OptionAffichage()  As typOptAff
Public DefOptAff As typOptAff


' Fonctions de  l'API Windows
'--------------------------------------
' Compilation conditionnelle Win32
#If Win32 Then
  Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
  
  Private Declare Function GetPrivateProfileInt Lib "Kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal n As Integer, ByVal lpFilename As String) As Integer
#End If
  
Public Sub Main()
Dim NomFich As String

GirationVersion = App.Major & "." & App.Minor
gtCoordActif = True

#If GIRATIONDEMO Then
  VersionDemo = True
#End If
gtAppliDistante = Left(App.Path, 2) = "\\"
 
Set gtFso = CreateObject("Scripting.FileSystemObject")
 
 GetCommandLine NomFich
  
  If FichierJournal Then
    numFichLog = FreeFile
    Open App.Path & "\giration.log" For Output As numFichLog
  End If
  
  On Error GoTo GestErr
  
'********************************
'test Protection
'********************************
  'Type de protection
  TYPPROTECTION = CPM
  ' Vérification de l'enregistrement
  If gtVersionDéveloppeur Then
    ' Affichage de la feuille principale
     MDIGiration.Show
  ElseIf ProtectCheck("its00+-k") = "its00+-k" Then
    ' Affichage de la feuille principale
     MDIGiration.Show
  Else 'la licence n'a pas été validée on ferme
     End
  End If
'********************************

 ' Initialisation de l'imprimante (NT ou non)
  Set gDlgPrint = MDIGiration.dlgPrint
  ' InitializePrinter différé dans
'  InitializePrinter
  ' Initialisations générales
  gtIniVar
   
  If ExistFich(NomFich) Then
    MDIGiration.dlgTrajectoire.FileName = NomFich
    creFille nouveau:=False
  End If
  
  Exit Sub
  
GestErr:

  ErrGeneral
    
  End Sub

'******************************************************************************************
' Initialisation des variables globales de GIRATION
'******************************************************************************************
Private Sub gtIniVar()
  Dim i As Integer
  
' Initialisation des noms de plans réservés
  For i = 0 To 7: GirPlanReserve(i) = "$$$GIR_" & CStr(i): Next

' AV : 01/02/2000 : les constantes ci-dessous ne doivent pas dépasser 11 caractères
' car elles servent à construire un nom de plan DXF (<32 car.) dont le préfixe est : GIR_<VEHICULE.nomCOURT>
' et <VEHICULE.nomCourt> est lui-même tronqué à 16 (cf VEHICULE.CLS)
  radicExport(0) = ID_ExportTracteur
  radicExport(1) = ID_ExportRemorque
  radicExport(2) = ID_ExportRemorque2
  radicExport(3) = ID_ExportRouesAvt
  radicExport(4) = ID_ExportRouesArr
  radicExport(5) = ID_ExportRouesRmq1
  radicExport(6) = ID_ExportRouesRmq2
  
  VersionDemoNom = "Véhicules v. DEMO" & Chr(26)
  If VersionDemo Then
    vehicVersion.Nom = VersionDemoNom
  Else
    vehicVersion.Nom = "Véhicules v. 1.0" & Chr(26)
  End If

' Définition des correspondances de couleur GIRATION-AutoCAD
  tableCouleur
  InvTableCouleur

' unités d'angle
  libUnite(DEGRE) = ID_Degres
  libUnite(GRADE) = ID_Grades
  eqvPI(DEGRE) = 180
  eqvPI(GRADE) = 200
  pi = 4 * Atn(1)
 
  grisPale = RGB(255, 255, 232)
  blanc = RGB(255, 255, 255)

  
' Fichiers GIRATION : biblio de véhicules
  GirationVeh = App.Path & "\" & GIRATIONVEHCOURT
  DefautRep(0) = "FDP"
  DefautRep(1) = "TRAJ"
  
  ' Lecture de la base de registres
  LireRegistry
  ' Lecture de la Bibliothèque de véhicules
  lireVeh
  
  gtBoutons(0) = "btnFleche"
  gtBoutons(1) = "btnXY"
  gtBoutons(2) = "btnDistance"
  gtBoutons(3) = "btnAngle"
  
End Sub

Public Sub GriserMenus(Etat As Integer)
Dim i%, PasAPasActif As Boolean

  With MDIGiration
    'Fichier Fermer
    .mnuFich(3).Enabled = Etat
    
    'Fichier Enregistrer la trajectoire
    .mnuFich(MNUENREGTRAJ).Enabled = Etat
    
    'Fichier Enregistrer la trajectoire sous
    .mnuFich(MNUENREGTRAJSOUS).Enabled = Etat
    
    'Fichier Importer une bibli de véhicules
    .mnuFich(MNUIMPORTBIB).Enabled = Not VersionDemo
    
    'Fichier Exporter une trajectoire
    .mnuFich(MNUEXPORTEPURE).Enabled = Etat
    
    'Fichier Exporter dans un fichier texte
    .mnuFich(MNUEXPORTCARAC).Enabled = Etat
    
    'Fichier Importer un fond de plan
    .mnuFich(MNUIMPORTFDP).Enabled = Etat
    
    'Fichier Enregistrer le fond de plan sous
    .mnuFich(MNUENREGFDP).Enabled = Etat
    
    'Fichier Imprimer
    .mnuFich(MNUIMPRIMER).Enabled = Etat And Printers.Count > 0
    
    'Outils Insérer un déplacement
    .mnuOutils(MNUINSERDEPL).Enabled = Etat
        
    'Outils Supprimer un déplacement
    .mnuOutils(MNUSUPPRDEPL).Enabled = Etat
        
    'Outils Virage
    .mnuOutils(MNUDECOMPVIRAGE).Enabled = Etat

    'Outils Pas à pas
    .mnuOutils(MNUPASAPAS).Enabled = Etat
    
    'Outils Surlargeur
    .mnuOutils(MNULARGEUR).Enabled = Etat
    
    'Outils Cadrage
    .mnuOutils(MNUCADRAGE).Enabled = Etat
     
    'Outils Rafraichir
    .mnuOutils(MNURAFRAICHIR).Enabled = Etat
     
    'Outils Options générales
'v3.2
    .mnuOutils(MNUOPTGENERAL).Enabled = True

    'Outils Options d'affichage
    .mnuOutils(MNUOPTAFFICHAGE).Enabled = Etat

    'Outils Mesures
    .mnuOutils(MNUMESURES).Enabled = True
    'Outils Angle Total
    .mnuOutils(MNUANGLETOTAL).Enabled = True

    If Etat Then
      ' Désactivation de certains items selon le contexte
      .mnuFich(MNUENREGTRAJ) = fCourante.Modif And Not VersionDemo                                              ' Enregistrer trajectoire
      .mnuFich(MNUENREGTRAJSOUS) = fCourante.Grille.MaxRows > 2 And Not VersionDemo              ' Enregistrer trajectoire sous
      .mnuFich(MNUENREGFDP).Enabled = fCourante.FDP <> "" And Not VersionDemo                                             'Enregistrer le fond de plan
      With OptionAffichage(NumFille)   ' Pas à pas si au moins une carrosserie à dessiner
        For i = 0 To fCourante.nbRemorque
          PasAPasActif = PasAPasActif Or .Carross(i).ADessiner
        Next
      End With
         ' Export et Pas à pas si au moins un déplacement (<==> ExportOK)
      .mnuFich(MNUEXPORTEPURE).Enabled = fCourante.ExportOK And Not VersionDemo
      .mnuOutils(MNUPASAPAS).Enabled = PasAPasActif And fCourante.ExportOK
      .mnuZoom(TOOL_ZAVANT).Enabled = fCourante.Milieux.Count > 1
    End If

    .tbrGiration.Buttons("btnSave").Enabled = .mnuFich(4)
    .tbrGiration.Buttons("btnPrint").Enabled = .mnuFich(MNUIMPRIMER)
    
    .tbrGiration.Buttons("btnZoom").Enabled = Etat
    .tbrGiration.Buttons("btnZAvant").Enabled = Etat
    .tbrGiration.Buttons("btnPAN").Enabled = Etat
    .tbrGiration.Buttons("btnFleche").Enabled = Etat
    .tbrGiration.Buttons("btnXY").Enabled = Etat
    .tbrGiration.Buttons("btnDistance").Enabled = Etat
    .tbrGiration.Buttons("btnAngle").Enabled = Etat
    If Etat Then .tbrGiration.Buttons("btnZAvant").Enabled = (fCourante.Milieux.Count > 1)
  
  End With

End Sub

'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder

'Public Sub lireProtect()

   ' Do While MyCCToujours() <> 0
   '   Select Case MsgBox("Protection retirée", vbRetryCancel + vbCritical, ID_GestionProtection)
   '   Case vbCancel
   '     gtProtectRetirée = True
   '     Exit Do
   '   End Select
   ' Loop
    
   ' If gtProtectRetirée Then Unload MDIGiration

'End Sub

Private Sub LireRegistry()

' Variant devant contenir un tableau à deux dimensions renvoyé par la fonction GetAllSettings
Dim MySettings As Variant
' Nombre entier correspondant à la valeur du compteur.
Dim intSettings As Integer
Dim Setting As Variant

Dim chaine As String
Dim RegistryAECrire As Boolean
Dim i%

  If IsNumeric("1.1") Then
    gtPtDecimal = Asc(".")  ' 46
  Else
    gtPtDecimal = Asc(",")  ' 44
  End If

                            ' Section Parametres
  With OptGen
    .unite = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:="Unite", Default:=DEGRE)

    chaine = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:="Vitesse max", Default:=15#)
  ' L'utilisateur a pu changer de config au niveau du séparateur décimal
    If IsNumeric(chaine) Then
      .VitMax = chaine
    Else
      .VitMax = substPtDecimalRegional(chaine, Regional:=True)
    End If

    chaine = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:="Vitesse defaut", Default:=5#)
  ' L'utilisateur a pu changer de config au niveau du séparateur décimal
    If IsNumeric(chaine) Then
      .VitDef = chaine
    Else
      .VitDef = substPtDecimalRegional(chaine, Regional:=True)
    End If

    .Epaisseur = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:="Epaisseur ligne", Default:=1)

  End With

  MDIGiration.mnuOutils(MNUMESURES).Checked = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:="Mesure", Default:=0)
  MDIGiration.mnuOutils(MNUANGLETOTAL).Checked = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:="TotalAngle", Default:=0)

  GirInitOk = True
  On Error GoTo GestErr
  
  LireRepertoires
  

If Not GirInitOk Then
  frmOptGen.Show vbModal
  If GirInitOk Then
    RegistryAECrire = True
  Else
    MsgBox IDm_InitInterrupt
    Unload MDIGiration
  End If
End If

' Mémorisation du chemin d'un import DXF pour l'import suivant dans la même session
gtRepertFDP = OptGen.Repert(0)
' Mémorisation du chemin de la trajectoire pour l''ouvrir' suivant dans la même session
MDIGiration.dlgTrajectoire.InitDir = OptGen.Repert(1)

'If maxKey = 2 Then
'  GirationVeh = OptGen.Repert(2) & "\" & GIRATIONVEHCOURT
'End If

' Extrait les paramètres.
MySettings = GetAllSettings(Appname:=App.Title, SECTION:="Recent Files")
If Not IsEmpty(MySettings) Then
  For intSettings = UBound(MySettings, 1) To 0 Step -1     ' On les lit  à l'envers, car MRUmenu les ajoute par décalage
    MRUmenu MySettings(intSettings, 1)
  Next
End If

                              ' Section Affichage
  With DefOptAff
  
    .Carross(0).ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Tracteur", Default:=1)
  
    .Carross(1).ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Remorque1", Default:=1)
  
    .Carross(2).ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Remorque2", Default:=1)
  
    .RouAvt.ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Roues avant", Default:=1)
  
    .RouArr(0).ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Roues arrière", Default:=1)
  
    .RouArr(1).ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Roues remorque1", Default:=1)
  
    .RouArr(2).ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Roues remorque2", Default:=1)
  
    .FDP.ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Fond de plan", Default:=1)
  
    .Graduations.ADessiner = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Graduations", Default:=1)
  
    .sens = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Marquage sens", Default:=0)
  
    .position = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Indication position", Default:=0)
    
    .pas(0) = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Increment1", Default:=2#)
     
    .pas(1) = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Increment2", Default:=5#)
    
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur1", Default:="0 0 255")  ' bleu
    .Carross(0).couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur2", Default:="0 255")  ' vert
    .Carross(1).couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur3", Default:="255")  ' rouge
    .Carross(2).couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur4", Default:="0")
    .RouAvt.couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur5", Default:="0")
    .RouArr(0).couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur6", Default:="0")
    .RouArr(1).couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur7", Default:="0")
    .RouArr(2).couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur8", Default:="-1")
    .FDP.couleur = recompRGB(chaine)
  
    chaine = GetSetting(Appname:=App.Title, SECTION:="Affichage", Key:="Couleur9", Default:="0 128 128")
    .Graduations.couleur = recompRGB(chaine)

    If RegistryAECrire Then EcrireRegistry
  End With


Exit Sub

GestErr:
  ErrGeneral
  
End Sub
  
Private Sub LireRepertoires()
Dim chaine As String
Dim Key(2) As String
Dim i As Integer

  Key(0) = "FDP"
  Key(1) = "Trajectoire"
  Key(2) = "Véhicules"

  ' Cahier des charges v3.2 : comme dans la v3.0 et 3.1, pas de msg d'erreur si la bibliothèque est vide
  chaine = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:=Key(2))
  If chaine = "" Then OptGen.Repert(2) = App.Path Else OptGen.Repert(2) = supprSlash(chaine)
  GirationVeh = gtFso.BuildPath(OptGen.Repert(2), GIRATIONVEHCOURT)

  GirInitOk = Not VerifDroitBibVeh(OptGen.Repert(2))
  
  For i = 0 To 1
    chaine = GetSetting(Appname:=App.Title, SECTION:="Parametres", Key:=Key(i))

    If Not ExistFich(chaine, vbDirectory) Then
      GirInitOk = False
      OptGen.Repert(i) = "C:\Mes Documents"   ' Modif provisoire v3.2 (18/05/2000)
   Else
      OptGen.Repert(i) = supprSlash(chaine)
    End If
    If ExistFich(OptGen.Repert(i), vbDirectory) Then OptGen.Repert(i) = gtFso.GetFolder(OptGen.Repert(i)).Path
  Next

End Sub

'***********************************************************************************
' Récupération d'un éventuel Giration.ini issu d'une version précédente de Giration
'***********************************************************************************
Private Function lireIni() As Boolean
' Lecture de Giration.ini
  Dim i%, n%, chaine As String
  
  ' Dimensionnement de la chaine réceptrice pour appel à GetProfileString
  chaine = String(256, Chr(0))
  
  GirationIni = App.Path & "\Giration.ini"
  
  If Not ExistFich(GirationIni) Then Exit Function
  
                            ' Section Parametres
  With OptGen
    n% = GetPrivateProfileString("Parametres", "FDP", "", chaine, 256, GirationIni)
    If n% = 0 Then Exit Function
    .Repert(0) = Left(chaine, n%)
    
    n% = GetPrivateProfileString("Parametres", "Trajectoire", "", chaine, 256, GirationIni)
    If n% = 0 Then Exit Function
    .Repert(1) = Left(chaine, n%)
   
    n% = GetPrivateProfileString("Parametres", "Unite", "Degré", chaine, 6, GirationIni)
    If suppCNull(chaine) = "Degré" Then .unite = 0 Else .unite = 1
    
    n% = GetPrivateProfileString("Parametres", "Vitesse max", "15", chaine, 5, GirationIni)
    .VitMax = valChaine(chaine)
    
    n% = GetPrivateProfileString("Parametres", "Vitesse defaut", "5", chaine, 5, GirationIni)
    .VitDef = valChaine(chaine)
  
    n% = GetPrivateProfileString("Parametres", "Epaisseur ligne", "1", chaine, 2, GirationIni)
    .Epaisseur = valChaine(chaine)

  End With

  
                              ' Section Affichage
  With DefOptAff
  
    .Carross(0).ADessiner = (GetPrivateProfileInt("Affichage", "Tracteur", 1, GirationIni) <> 0)
  
    .Carross(1).ADessiner = (GetPrivateProfileInt("Affichage", "Remorque1", 1, GirationIni) <> 0)
  
    .Carross(2).ADessiner = (GetPrivateProfileInt("Affichage", "Remorque2", 1, GirationIni) <> 0)
  
    .RouAvt.ADessiner = (GetPrivateProfileInt("Affichage", "Roues avant", 1, GirationIni) <> 0)
  
    .RouArr(0).ADessiner = (GetPrivateProfileInt("Affichage", "Roues arrière", 1, GirationIni) <> 0)
  
    .RouArr(1).ADessiner = (GetPrivateProfileInt("Affichage", "Roues remorque1", 1, GirationIni) <> 0)
  
    .RouArr(2).ADessiner = (GetPrivateProfileInt("Affichage", "Roues remorque2", 1, GirationIni) <> 0)
  
    .FDP.ADessiner = (GetPrivateProfileInt("Affichage", "Fond de plan", 1, GirationIni) <> 0)
  
    .Graduations.ADessiner = (GetPrivateProfileInt("Affichage", "Graduations", 1, GirationIni) <> 0)
  
    .sens = (GetPrivateProfileInt("Affichage", "Marquage sens", 0, GirationIni) <> 0)
  
    n% = GetPrivateProfileString("Affichage", "Increment1", "2.0", chaine, 5, GirationIni)
    .pas(0) = valChaine(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Increment2", "5", chaine, 3, GirationIni)
    .pas(1) = valChaine(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur1", "0 0 255", chaine, 20, GirationIni)   ' bleu
    .Carross(0).couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur2", "0 255", chaine, 20, GirationIni)       ' vert
    .Carross(1).couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur3", "255", chaine, 20, GirationIni)           ' rouge
    .Carross(2).couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur4", "0", chaine, 20, GirationIni)
    .RouAvt.couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur5", "0", chaine, 20, GirationIni)
    .RouArr(0).couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur6", "0", chaine, 20, GirationIni)
    .RouArr(1).couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur7", "0", chaine, 20, GirationIni)
    .RouArr(2).couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur8", "-1", chaine, 20, GirationIni)                    ' -1 =  Couleurs originales du fond de plan
    .FDP.couleur = recompRGB(chaine)
  
    n% = GetPrivateProfileString("Affichage", "Couleur9", "0 128 128", chaine, 20, GirationIni)       ' vert cacadois
    .Graduations.couleur = recompRGB(chaine)

  End With

                            ' Section Trajectoires
  For i = 4 To 1 Step -1   ' On les lit  à l'envers, car MRUmenu les ajoute par décalage
    n% = GetPrivateProfileString("Trajectoires", "Trajectoire" & CStr(i), "", chaine, 255, GirationIni)
    MRUmenu suppCNull(chaine)
  Next

  EcrireRegistry
  lireIni = True

End Function

Public Function VerifDroitBibVeh(ByVal RepVehicule As String) As Boolean
  
  If Not ExistFich(RepVehicule, vbDirectory) Then
    VerifDroitBibVeh = True
  Else
    On Error GoTo GestErr
    Dim numFich As Integer
    numFich = FreeFile
    Open gtFso.BuildPath(RepVehicule, GIRATIONVEHCOURT) For Random Access Read As numFich
    Close numFich
  End If
  
  Exit Function

GestErr:
  If Err = 75 Then
    MsgBox IDm_DroitsBibVeh
    VerifDroitBibVeh = True
  Else
    ErrGeneral
  End If

End Function

Public Sub lireVeh(Optional ByVal Relecture As Boolean)
Dim i%, numFich%, Version As StructVéhicule, VehicTab As StructVéhicule
Dim nbVehicul As Integer, Ajouter As Boolean
Dim lgEnregVehic As Integer
Dim Vehic As VEHICULE

  If VersionDemo And Relecture Then Exit Sub

  For Each Vehic In ColVéhicules
    ColVéhicules.Remove UCase(Vehic.Nom)
  Next

  On Error GoTo GestErr
    ' le 1er enregistrement contient le  numéro de version de fichiers de la bibliothèque (préserve l'évolution)
  lgEnregVehic = Len(vehicVersion)
  nbVehicul = FileLen(GirationVeh) / lgEnregVehic - 1
  
    If VersionDemo And nbVehicul <> 3 Then
      MsgBox IDm_IncompatiblBibvehVersiondemo
      Exit Sub
    End If
    
    numFich = FreeFile
    Open GirationVeh For Random Access Read As numFich Len = lgEnregVehic
    Get #numFich, 1, Version
    If VersionDemo And Version.Nom <> VersionDemoNom Then
      MsgBox IDm_IncompatiblBibvehVersiondemo
      Close numFich
      Exit Sub
    End If
    
    If Version.Nom = VersionDemoNom Then vehicVersion.Nom = VersionDemoNom
    
    
    For i = 1 To nbVehicul
      Get #numFich, i + 1, VehicTab
        
        If Version.Nom = VersionDemoNom Then
        ' décryptage des noms de véhicule d'une éventuelle version de démo
          Select Case decrypt(Trim(VehicTab.Nom))
          Case "Simple"
            Ajouter = True
            VehicTab.Nom = ID_Simple  'frmBibVéhicule.optTypVeh(0).Caption
'            Vehictab.Nom = "Simple"
          Case "Articulé"
            Ajouter = True
'            Vehictab.Nom = "Articulé"
            VehicTab.Nom = ID_Articulé    'frmBibVéhicule.optTypVeh(1).Caption
          Case "BiArticulé"
            Ajouter = True
'            Vehictab.Nom = "BiArticulé"
            VehicTab.Nom = ID_BiArticulé  'frmBibVéhicule.optTypVeh(2).Caption
          Case Else
            Ajouter = False
          End Select
          
        Else  ' Version normale
          Ajouter = True
        End If
        
      If Ajouter Then InserVéhicule VehicTab
    
    Next
    
    Close numFich
    
  Exit Sub
  
GestErr:
  If Err = 53 Then
  ElseIf Err = 75 Then
    MsgBox "Droits d'accès à la bibliothèque de véhicules insuffisants"
  Else
    ErrGeneral  ' 53 : Fichier introuvable (bibliothèque renommée ou déplacée)
  End If
 
End Sub


'******************************************************************************
' Retourne si une feuille est chargée
'******************************************************************************
Public Function EstChargée(ByVal FeuilleCherchée As Form) As Boolean
Dim Feuille As Form
  For Each Feuille In Forms
    If Feuille Is FeuilleCherchée Then EstChargée = True: Exit For
  Next
End Function

'******************************************************************************
' Détermination de la position d'insertion d'un véhicule dans la collection
'******************************************************************************
Public Function InserVéhicule(VehicTab As StructVéhicule) As VEHICULE
Dim i As Integer

On Error GoTo GestErr

  With VehicTab
    If ColVéhicules.Count = 0 Then
    '1er véhicule : forcément au début
      Set InserVéhicule = AjoutVéhicule(VehicTab)

    Else
        If StrComp(.Nom, ColVéhicules(1).Nom, vbTextCompare) = -1 Then
        ' le véhicule est inséré en 1ère position
        Set InserVéhicule = AjoutVéhicule(VehicTab, 1)
      ElseIf StrComp(ColVéhicules(ColVéhicules.Count).Nom, .Nom, vbTextCompare) = -1 Then
        ' le véhicule est ajouté en fin de liste
        Set InserVéhicule = AjoutVéhicule(VehicTab)
      Else
        'Rechercher en quelle position l'insérer
        For i = 2 To ColVéhicules.Count
          If StrComp(.Nom, ColVéhicules(i).Nom, vbTextCompare) = -1 Then
            ' le véhicule est inséré à sa place
            Set InserVéhicule = AjoutVéhicule(VehicTab, i)
            Exit For
          End If
        Next
        
        ' le véhicule doit être ajouté en fin de liste
        If InserVéhicule Is Nothing Then Set InserVéhicule = AjoutVéhicule(VehicTab)
      End If
    End If
  End With

Exit Function
GestErr:
  If Err <> ErrVéhicule Then
    ErreurNonFatale "InserVéhicule"
  End If
  Err.Raise ErrVéhicule
End Function


'******************************************************************************
' Ajout du véhicule dans la collection à la position voulue
'******************************************************************************
Private Function AjoutVéhicule(w As StructVéhicule, Optional position As Integer) As VEHICULE
Dim VehicNew As New VEHICULE
Dim numRemorque As Integer

On Error GoTo GestErr

  With VehicNew
    .Nom = Trim(w.Nom)
    .Protégé = w.Protégé
    .nbRemorque = w.nbRemorque
    For numRemorque = 0 To .nbRemorque
      .PAFav(numRemorque) = w.PAFav(numRemorque)
      .e(numRemorque) = w.e(numRemorque)
      .PAFar(numRemorque) = w.PAFar(numRemorque)
      .l(numRemorque) = w.l(numRemorque)
      .LEar(numRemorque) = w.LEar(numRemorque)
      If numRemorque > 0 Then
        .Dpat(numRemorque - 1) = w.Dpat(numRemorque - 1)
      End If
    Next
    .LEav = w.LEav
    .alphaVehMax = w.alphaVehMax
    .aVehMax = w.aVehMax
    .RVehMin = w.RVehMin
  End With
    
'La clé du véhciule est son nom en majuscules
  If position = 0 Then   'ajout en fin de liste
    ColVéhicules.Add VehicNew, UCase(VehicNew.Nom)
  Else                                             ' insertion dans la liste
    ColVéhicules.Add VehicNew, UCase(VehicNew.Nom), before:=position
  End If
  
  Set AjoutVéhicule = VehicNew

Exit Function

GestErr:
  If Err <> ErrVéhicule Then
    ErreurNonFatale "InserVéhicule"
  End If
  Err.Raise ErrVéhicule
End Function


'******************************************************************************
' Mise au format StructVéhicule de l'objet Véhicule en vue de son écriture dans GIRATION.VEH
'******************************************************************************
Public Sub ConvObjetStruct(VehicNew As VEHICULE, w As StructVéhicule)
Dim numRemorque As Integer
    
  With VehicNew
    If vehicVersion.Nom = VersionDemoNom Then
      w.Nom = Trim(crypt(.Nom))
    Else
      w.Nom = Trim(.Nom)
    End If
    w.Protégé = .Protégé
    w.nbRemorque = .nbRemorque
    For numRemorque = 0 To .nbRemorque
      w.PAFav(numRemorque) = .PAFav(numRemorque)
      w.e(numRemorque) = .e(numRemorque)
      w.PAFar(numRemorque) = .PAFar(numRemorque)
      w.l(numRemorque) = .l(numRemorque)
      w.LEar(numRemorque) = .LEar(numRemorque)
      If numRemorque > 0 Then
        w.Dpat(numRemorque - 1) = .Dpat(numRemorque - 1)
      End If
    Next
    w.LEav = .LEav
    w.alphaVehMax = .alphaVehMax
    w.aVehMax = .aVehMax
    w.RVehMin = .RVehMin
  End With

End Sub

Private Function decrypt(v As String) As String
' DéCryptage pour la version Démo
Dim i%, c As String

  For i = 1 To Len(v)
    c = Mid(v, i, 1)
    If c <> " " Then c = Chr(Asc(c) + 64)
    decrypt = decrypt + c
  Next
  
End Function

Private Function crypt(v As String) As String
' Cryptage pour la version Démo
Dim i As Integer, c As String

  For i = 1 To Len(v)
    c = Mid(v, i, 1)
    If c <> " " Then c = Chr(Asc(c) - 64)
    crypt = crypt + c
  Next
  
End Function

Private Sub EcrireRegistry()
Dim res As Integer

  With DefOptAff
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Tracteur", Setting:=valBool(.Carross(0).ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Remorque1", Setting:=valBool(.Carross(1).ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Remorque2", Setting:=valBool(.Carross(2).ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues avant", Setting:=valBool(.RouAvt.ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues arrière", Setting:=valBool(.RouArr(0).ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues remorque1", Setting:=valBool(.RouArr(1).ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues remorque2", Setting:=valBool(.RouArr(2).ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Fond de plan", Setting:=valBool(.FDP.ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Graduations", Setting:=valBool(.Graduations.ADessiner)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Marquage sens", Setting:=valBool(.sens)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Indication position", Setting:=valBool(.position)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Increment1", Setting:=.pas(0)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Increment2", Setting:=.pas(1)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur1", Setting:=decompRGB(.Carross(0).couleur)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur2", Setting:=decompRGB(.Carross(1).couleur)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur3", Setting:=decompRGB(.Carross(2).couleur)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur4", Setting:=decompRGB(.RouAvt.couleur)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur5", Setting:=decompRGB(.RouArr(0).couleur)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur6", Setting:=decompRGB(.RouArr(1).couleur)
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur7", Setting:=decompRGB(.RouArr(2).couleur)
  
    If .FDP.couleur = -1 Then
      SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur8", Setting:=-1
    Else
      SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur8", Setting:=decompRGB(.FDP.couleur)
    End If
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur9", Setting:=decompRGB(.Graduations.couleur)
  End With
 
End Sub
  
Private Function valBool(v As Boolean) As String
' Transforme une valeur booléenne en chaine : "1" (True)  ou "0" (False)

  If v Then valBool = "1" Else valBool = "0"
  
End Function

Public Sub creFille(ByVal nouveau As Boolean, Optional ByVal Nom As String)
' En création, nom représente le nom de l'éventuel Fond de Plan saisi dans frmTrajPar
' En ouverture, nom représente le nom du  fichier .GIR ou .EGT de la trajectoire (supprimé v3.2 : 27/11/2000)
Dim i%, Cancel As Boolean
Dim SauveCourante As frmTraject

On Error GoTo GestErr

'====================================================================================================
'Remarque importante
' Numfille est stockée définitivement dans la propriété Tag de l'instance frmTraject
' La fermeture d'une feuille ne modifie pas cette valeur pour les autres feuilles ni celle de NbreFilles
'====================================================================================================
  
  ReDim Preserve OptionAffichage(NbreFilles)
  
  NumFille = NbreFilles
  
  If gtFeuillesTraject.Count > 0 Then Set SauveCourante = fCourante
  If NbreFilles = 0 Then
   Set fCourante = frmTraject
  Else
    Set fCourante = New frmTraject
  End If
  
  fCourante.nouveau = nouveau
  If nouveau Then
    If Nom = "" Then
      If FichierJournal Then Write #numFichLog, "Nouvelle trajectoire sans FDP"
    Else
      If FichierJournal Then Write #numFichLog, "Nouvelle trajectoire avec FDP : " & Nom
    End If
          
    fCourante.Caption = Majus(ID_Trajectoire) & CStr(NbreFilles + 1)    ' Déclenche Form_Load
  Else
    With MDIGiration.dlgTrajectoire
      Nom = .FileName
      If Extension(Nom, vbLowerCase) = "gir" Then
        fCourante.Caption = nomCourt(Nom)
        If fCourante.Tag = "" Then Err.Raise 1001 'le chargement a échoué
        fCourante.dlgTrajectoire.FileName = Nom
        MRUmenu Nom     'MRUmenu est défini dans GIRATION.BAS
      Else
        fCourante.Caption = Majus(ID_Trajectoire) & CStr(NbreFilles + 1)
      End If
      .FileName = ""  ' Réinitialisation pour l'ouverture suivante
    End With
    
  End If
  
  fCourante.Show
  'L'instruction suivante a été rajoutée le 11/05/2000 (AV) : Le Form_Activate ne suffisant plus(??) à empêcher les affichages successifs des controles
  DoEvents
  
  ActivMenu Activation:=True
  fCourante.Form_Activate
  ActivMenu Activation:=False
  
  fCourante.CreCarross
  
  If nouveau Then fCourante.FDP = Nom
  If fCourante.FDP = "" Then     ' pas de FDP
    Cancel = True
  Else
    If Not nouveau Then fCourante.Grille.Visible = False  ' sinon, la boite DéfileDXF se cache tout-de-suite
    lecFDP fCourante.FDP, Cancel     ' lecFDP recalcule les nouveaux pmin et pmax de la fenêtre trajectoire
  End If
  
  verifRetaillage Ouverture:=True  ' si le retaillage a eu lieu dans  une autre fenêtre, il faut réinitialiser l'origine
  
  If Not nouveau Then
    ToutVoir
  Else
    If Cancel Then ' pas de FDP ou lecture abortée
      fCourante.FDP = ""
      gtFacteurZoom = 50
      fCourante.Echelles.Add gtFacteurZoom, gtFacteurZoom
      fCourante.Milieux.Add 0, 0
      Set gtMil = fCourante.Milieux(1)   'fCourante.Milieux.count)
      
      ParamEcranZoom
    Else
      CalcEchelle
    End If
    OrigineParDefaut
  End If
  
  If fCourante.FDP <> "" Then
'Dessin du fond de plan
    desFDP
    If Extension(fCourante.FDP, vbLowerCase) <> "fdp" Then
      fCourante.dlgFDP.FileName = ""   ' ainsi, la fonction Enregistrer FDP demandera un nom
    Else
      fCourante.dlgFDP.FileName = fCourante.FDP
    End If
  End If
  
  Graduer
  
  ActivMenu Activation:=True
  GriserMenus Etat:=True
  
  If nouveau Then
    DesCarrossPourOrigine
    fCourante.BasculeGrilleOrigine False
  Else
    fCourante.BasculeGrilleOrigine True
    fCourante.dessiner "TOUT"
    fCourante.autorRayon
  End If
  
  NbreFilles = NbreFilles + 1
  gtFeuillesTraject.Add fCourante, fCourante.Tag
  fCourante.SetFocus
   
Exit Sub

GestErr:
  If Err = 1001 Then ' Le chargement de la feuille a échoué
    Unload fCourante
    ActivMenu Activation:=True
    If Not SauveCourante Is Nothing Then SauveCourante.Form_Activate
  Else
  
    ErrGeneral "frmTraject.creFille"
  End If

End Sub


Private Sub MRUmenu(ByVal NomFich As String)
' Affichage dans le menu des derniers fichiers utilisés (MRU)
Dim i As Integer, j As Integer
Const MAXFICH = 4

  
  If Not ExistFich(NomFich) Then Exit Sub
  
  For i = 0 To nbfichMenu - 1
    If UCase(NomFich) = UCase(MRUFichiers(i)) Then
      ' Le fichier est déjà dans la liste, on va le remettre en tête
      For j = i To nbfichMenu - 2
      ' Resserrement du tableau en supprimant le fichier
        MRUFichiers(j) = MRUFichiers(j + 1)
      Next
      nbfichMenu = nbfichMenu - 1
    End If
  Next

  If nbfichMenu < MAXFICH Then    ' Debug: cette valeur pourra être paramétrée (n'apparait qu'ici), mais sans dépasser la taille de mnuSelect
    MDIGiration.mnuSelect(nbfichMenu).Visible = True
    If nbfichMenu = 0 Then MDIGiration.mnuSelect(4).Visible = True 'Création de la ligne de séparation
    ReDim Preserve MRUFichiers(nbfichMenu)
    nbfichMenu = nbfichMenu + 1
  End If
  
  For i = nbfichMenu - 1 To 1 Step -1
    ' Décalage des fichiers pour insérer le nouveau en tête
    MRUFichiers(i) = MRUFichiers(i - 1)
    MDIGiration.mnuSelect(i).Caption = "&" & CStr(i + 1) & " " & MRUFichiers(i)
  Next
  MRUFichiers(0) = NomFich
  MDIGiration.mnuSelect(0).Caption = "&1 " & NomFich
  
End Sub

Public Sub Ouvrir()
Dim Cancel As Boolean

On Error GoTo ErrHandler
  
  With MDIGiration.dlgTrajectoire
    .InitDir = OptGen.Repert(1)
   .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    .ShowOpen
    
    If Not Cancel Then
'              .InitDir = extraiRep(.FileName)
      If FichierJournal Then Write #numFichLog, "Ouvrir " & .FileName
      creFille nouveau:=False     ' False= pas nouveau
      If FichierJournal Then Write #numFichLog, "Ouvrir OK"
    Else
      Screen.ActiveForm.SetFocus
    End If
    
  End With
          
Exit Sub

ErrHandler:   ' L'utilisateur a fait 'Annuler
  If Err = cdlCancel Then
    Cancel = True
    Resume Next
  Else
    ErrGeneral
  End If

End Sub

Public Sub Enregistrer(Feuille As Form, ByRef Cancel As Integer)
Dim NomFich As String

  If Feuille.numFichProtégé <> 0 Then Close Feuille.numFichProtégé
  
  NomFich = Feuille.dlgTrajectoire.FileName
  If NomFich <> "" And Not FichierProtégé(NomFich, , NomFich) Then
    Feuille.ecrire NomFich
  Else
    If Feuille.numFichProtégé <> 0 Then ' Réactivation de la protection
      Feuille.numFichProtégé = FreeFile
      Open NomFich For Random Lock Read As Feuille.numFichProtégé
    End If
    EnregSous Feuille, Cancel
  End If
  
End Sub

Public Sub EnregSous(Feuille As Form, ByRef Cancel As Integer)
  
  On Error GoTo ErrHandler
  
  With Feuille.dlgTrajectoire
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    If .InitDir = "" Then .InitDir = OptGen.Repert(1)
    .ShowSave
    If Not Cancel Then
      .InitDir = extraiRep(.FileName)
      If FichierProtégé(.FileName, , .FileName) Then
        Cancel = True
      Else
        Feuille.ecrire .FileName
        Feuille.Caption = gtFso.GetBaseName(.FileTitle)
        MRUmenu Feuille.dlgTrajectoire.FileName
      End If
    End If
  End With
  
  Exit Sub
  
ErrHandler:   ' L'utilisateur a fait 'Annuler
  Cancel = True
  Resume Next

End Sub

Public Sub EnregFDPSous(Feuille As Form, ByRef Cancel As Integer)
  
  On Error GoTo ErrHandler
  
  With Feuille.dlgFDP
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    .DialogTitle = ID_EnregFDP
    .Filter = ID_FilterFDP
    .InitDir = OptGen.Repert(0)   ' Répertoire par défaut des fonds de plan
    .ShowSave
    DoEvents
    If Not Cancel Then
      .InitDir = extraiRep(.FileName)
      If FichierProtégé(.FileName) Then Cancel = True: Exit Sub
    If FichierJournal Then Write #numFichLog, "Enregistrement du FDP " & .FileName
      EcrDXF.EcrDXF .FileName
      fCourante.FDP = .FileName
    If FichierJournal Then Write #numFichLog, "Enregistrement du FDP OK"
      
    Else                     ' Not Cancel
      Feuille.SetFocus
    End If
  End With

  Exit Sub
  
ErrHandler:   ' L'utilisateur a fait 'Annuler
  Cancel = True
  Resume Next
    
End Sub

Public Sub Import(Optional nomseul As Variant)
Dim Cancel As Boolean, origineVisible As Boolean
' nomseul indique que la lecture effective du fond de plan sera effectuée plus tard

  On Error GoTo ErrHandler
  
  With fCourante.dlgFDP
   .InitDir = gtRepertFDP
   .InitDir = OptGen.Repert(0)
   .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
   .DialogTitle = ID_ImportFDP
   .Filter = ID_FilterFDPGlobal
   .ShowOpen
   
    If Not Cancel Then
'       gtRepertFDP = extraiRep(.FileName)
      If IsMissing(nomseul) Then
        If FichierJournal Then Write #numFichLog, "Import FDP " & .FileName
        origineVisible = fCourante.fraOrigine.Visible
        fCourante.fraOrigine.Visible = False
        fCourante.Grille.Visible = False
        lecFDP .FileName, Cancel
        If origineVisible Then
          fCourante.fraOrigine.Visible = True
        Else
          fCourante.Grille.Visible = True
        End If
        If fCourante.Grille.MaxRows > 2 Or Cancel Then
        Else
          CalcEchelle
          OrigineParDefaut
        End If
        Rafraichir
      End If
      
      If Cancel Then       ' lecFDP a pu conduire à abandonner la lecture : Cancel=true
        If FichierJournal Then Write #numFichLog, "Import interrompu"
        fCourante.FDP = ""
      Else
        If FichierJournal Then Write #numFichLog, "Import OK"
        fCourante.FDP = .FileName   ' attention à l'écrasement (voir ci-dessus)
        If Extension(.FileName, vbLowerCase) <> "fdp" Then
          .FileName = ""   ' ainsi, la fonction Enregistrer FDP demandera un nom
        End If
      End If
      
      GriserMenus Etat:=True
      
    Else
      fCourante.SetFocus
    End If               ' Not Cancel
  End With
  
Exit Sub

ErrHandler:      ' L'utilisateur a fait 'Annuler
  Cancel = True
  Resume Next

End Sub

Public Sub Export()
Dim i%
Dim nomBloc As New BLOC, objPt As PT
'Dim objetPolyligne As POLYLINE
Dim Cancel As Boolean
Dim objPlan As PLAN
Dim numRemorque As Integer
Dim pLine As POLYLINE

  
  With fCourante.dlgExport
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    If .InitDir = "" Then .InitDir = App.Path       ' pas de répertoire par défaut prévu  : répertoire courant
  
  ' Activer un gestionnaire d'erreur particulier pour ShowSave
  On Error GoTo ErrHandler
    .ShowSave
  On Error GoTo 0
    
    If Not Cancel Then
      If FichierProtégé(.FileName) Then Exit Sub
      If FichierJournal Then Write #numFichLog, "Export dans " & .FileName
      flagExport = True
      
      ' Création des plans
      Set gtPlans = ExportPlans
      For i = 0 To 6: fCourante.crePlanReserve (i): Next
          
       ' Création des blocs de carrosserie
      For numRemorque = 0 To fCourante.nbRemorque
        If OptionAffichage(NumFille).Carross(numRemorque).ADessiner Then
          Set pLine = New POLYLINE
          With pLine
            .nomPlan = PlanExport(numRemorque)
            .ferme = True
          End With
          
          BlocExport(numRemorque) = "GIR_" & fCourante.Vehic.nomCourt & radicExport(numRemorque)
          Set nomBloc = fCourante.GraphExport.gphBlocs.Add(BlocExport(numRemorque))
          nomBloc.blkPolylines.AddPline pLine
          'Dupliquer la polyligne de carrosserie dans le bloc créé pour l'export
          For Each objPt In fCourante.Carross.gphBlocs(numRemorque + 1).blkPolylines(1).Pts
            pLine.Pts.Add objPt.X, objPt.Y
          Next
        End If
      Next

          ' Création du DXF
      fCourante.dessiner "TOUT"     ' fCourante.Dessiner appelle desCarross qui utilise le flag EXPORT pour créer des inserts au lieu de les dessiner
      
      If FichierJournal Then Write #numFichLog, "EcrDXF"
      
      EcrDXF.EcrDXF .FileName
    If FichierJournal Then Write #numFichLog, "FinEcrDXF"
      
          ' Suppression des plans
      For Each objPlan In ExportPlans
        gtPlans.Remove objPlan.name
      Next
                
          ' réinitialisation des objets graphiques exportés

      With fCourante.GraphExport
        Dim objBloc As BLOC
        For Each objBloc In .gphBlocs
          .gphBlocs.Remove objBloc.name
        Next
        While .gphInserts.Count <> 0 ' inserts de carrosseries
          .gphInserts.Remove 1
        Wend
        While .gphPolylines.Count <> 0   'roues
          .gphPolylines.Remove 1
        Wend
      End With

          ' Mémorisation pour une prochaine fois
      fCourante.Export = .FileName

      flagExport = False

      Set gtPlans = fCourante.colPlans
      If FichierJournal Then Write #numFichLog, "Export OK"

      Else
        fCourante.SetFocus
      End If    ' Not Cancel
  End With

  Exit Sub
  
ErrHandler:   ' L'utilisateur a fait 'Annuler
  Cancel = True
  Resume Next

End Sub

Public Sub ErrGeneral(Optional ByVal Fonction As String)
Dim message As String
    
    message = IDm_ErrFatale & CStr(Err) & vbCrLf & Err.Description & vbCrLf
    If Fonction <> "" Then message = message & vbCrLf & "Fonction " & Fonction & vbCrLf
    MsgBox message, vbOKOnly + vbCritical
    ErreurGenerale = True
    Unload MDIGiration

End Sub

Public Sub ErrLocal(Optional ByVal Feuille As Form, Optional ByVal Procédure As String)

    MsgBox IDm_Erreur & " " & CStr(Err) & vbCrLf & Err.Description & vbCrLf & Procédure, vbOKOnly + vbCritical
    If Not Feuille Is Nothing Then Unload Feuille

End Sub

'*************************************************************************************
' Décryptage de la ligne de commande
'*************************************************************************************
Private Function GetCommandLine(ByRef NomFich As String, Optional ByVal MaxArgs As Integer = 10)
' Déclare les variables.
  Dim c As String * 1
  Dim CmdLine As String
  Dim CmdLnLen As Integer
  Dim InArg As Boolean
  Dim i As Integer, NumArgs As Integer
  Dim ArgArray() As String
  Dim GuillemetOuvert As Boolean
  
' Définit un tableau au format approprié.
  ReDim ArgArray(MaxArgs - 1)
  NumArgs = 0: InArg = False
' Récupère les arguments de ligne de commande.
  CmdLine = Command()
  CmdLnLen = Len(CmdLine)
  
  If ExistFich(CmdLine) Then NomFich = CmdLine:  Exit Function
  
' Analyse de la ligne de commande caractère par caractère.
  For i = 1 To CmdLnLen
    c = Mid(CmdLine, i, 1)
' Analyse de caractères d'espacement ou de tabulations.
    If (c <> " " And c <> vbTab) Or GuillemetOuvert Then
      ' Ni espace ni tabulation.
    ' Vérifie une éventuelle présence dans l'argument.
      If Not InArg Then
        ' Le nouvel argument commence.
        ' Vérifie si les arguments ne sont pas trop nombreux.
        If NumArgs = MaxArgs Then Exit For
        NumArgs = NumArgs + 1
        InArg = True
      End If
      ' Concatène un caractère à l'argument courant.
      If c <> Chr(34) Then
        ArgArray(NumArgs - 1) = ArgArray(NumArgs - 1) & c
      Else
        GuillemetOuvert = Not GuillemetOuvert
      End If
    Else
      ' Recherche un espace ou une tabulation.
      ' L'indicateur InArg prend la valeur False.
      InArg = False
      If c = Chr(34) Then GuillemetOuvert = Not GuillemetOuvert
    End If
  Next i

' Redimensionne le tableau pour qu'il puisse
' juste contenir les arguments.
  If NumArgs > 0 Then ReDim Preserve ArgArray(0 To NumArgs - 1)

' Analyse logique des paramètres de la ligne de commande
  For i = 0 To NumArgs - 1
    If UCase(ArgArray(i)) = "/JOURNAL" Then
      FichierJournal = True
    ElseIf UCase(ArgArray(i)) = "/DEMO" Then
      VersionDemo = True
    ElseIf UCase(ArgArray(i)) = "/V" Then
      ' Seul le CERTU est autorisé à accéder au champ Verrouillage du véhicule
      If i < NumArgs And Not VersionDemo And UCase(ArgArray(i + 1)) = "CERTU" Then VerrouVéhicule = True
    ElseIf UCase(ArgArray(i)) = "/D" Then
      If i < NumArgs And UCase(ArgArray(i + 1)) = Chr(67) + Chr(69) & Chr(84) & Chr(69) & Chr(52) & Chr(52) Then
        gtVersionDéveloppeur = True
      End If
    ElseIf InStr(LCase(ArgArray(i)), ".gir") <> 0 Then
      NomFich = ArgArray(i)
    End If
  Next

' Renvoie le tableau dans le nom de fonction.
  GetCommandLine = ArgArray()
  
End Function


