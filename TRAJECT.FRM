VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmTraject 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Trajectoire"
   ClientHeight    =   7935
   ClientLeft      =   465
   ClientTop       =   585
   ClientWidth     =   12420
   Icon            =   "TRAJECT.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "TRAJECT.frx":030A
   ScaleHeight     =   7935
   ScaleWidth      =   12420
   Begin FPSpread.vaSpread Grille 
      Height          =   1932
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   8304
      _Version        =   131077
      _ExtentX        =   14647
      _ExtentY        =   3408
      _StockProps     =   64
      ButtonDrawMode  =   1
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      RestrictRows    =   -1  'True
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "TRAJECT.frx":042F
      UserResize      =   1
      VirtualOverlap  =   20
      VirtualRows     =   20
      VisibleCols     =   5
      VisibleRows     =   20
      VScrollSpecialType=   3
   End
   Begin VB.CommandButton cmdZMultiple 
      Caption         =   "Zoom multiple"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10440
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdPAN 
      Caption         =   "PAN"
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdZAvant 
      Caption         =   "Z avant"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "Zoom"
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdZArriere 
      Caption         =   "Zoom Arrière"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picTrajectoire 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FontTransparent =   0   'False
      Height          =   3996
      Left            =   120
      ScaleHeight     =   3930
      ScaleWidth      =   10650
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   10716
      Begin VB.Line linAngle 
         Visible         =   0   'False
         X1              =   5160
         X2              =   5280
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line linMarque 
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   2
         Index           =   1
         Visible         =   0   'False
         X1              =   8400
         X2              =   8400
         Y1              =   2040
         Y2              =   1320
      End
      Begin VB.Line linMarque 
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   8640
         X2              =   8640
         Y1              =   2040
         Y2              =   1320
      End
      Begin VB.Line linCarrosserie 
         Index           =   0
         Visible         =   0   'False
         X1              =   960
         X2              =   3120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Shape shpPcarac 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   30
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Shape shpPcarac 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   30
         Index           =   1
         Left            =   600
         Shape           =   3  'Circle
         Top             =   2520
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Shape shpPcarac 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   30
         Index           =   0
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   2520
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Shape shpZoom 
         Height          =   495
         Left            =   3600
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line linPan 
         Visible         =   0   'False
         X1              =   1920
         X2              =   2760
         Y1              =   2880
         Y2              =   2760
      End
   End
   Begin MSComDlg.CommonDialog dlgExport 
      Left            =   9360
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "dxf"
      DialogTitle     =   "Exporter l'épure sous"
      Filter          =   "Format DXF (*.dxf)|*.dxf"
   End
   Begin MSComDlg.CommonDialog dlgFDP 
      Left            =   8760
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "fdp"
      DialogTitle     =   "Enregistrer le fond de plan sous"
      MaxFileSize     =   2000
   End
   Begin MSComDlg.CommonDialog dlgTrajectoire 
      Left            =   8160
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "gir"
      Filter          =   "Trajectoire (*.gir)|*.gir"
      MaxFileSize     =   2000
   End
   Begin VB.PictureBox picBoutons 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   8640
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   12
      Top             =   5280
      Width           =   2775
      Begin VB.Frame fraAngleTotal 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label lblAngleTotal 
            Caption         =   "Angle Total"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdVéhicule 
         Height          =   615
         Left            =   1440
         Picture         =   "TRAJECT.frx":101E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Véhicule"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdBoutonOrigine 
         Height          =   615
         Left            =   360
         Picture         =   "TRAJECT.frx":1143
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Origine"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblCaracVeh 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   1440
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog dlgExportTexte 
      Left            =   9960
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "dxf"
      DialogTitle     =   "Exporter l'épure sous"
      Filter          =   "Texte(*.txt)|*.txt"
   End
   Begin VB.Frame fraOrigine 
      Caption         =   " Origine"
      Height          =   1695
      Left            =   720
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdSaisieOri 
         Height          =   375
         Index           =   1
         Left            =   4680
         Picture         =   "TRAJECT.frx":1559
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaisieOri 
         Height          =   375
         Index           =   0
         Left            =   3240
         Picture         =   "TRAJECT.frx":16C2
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtDistance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuler"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdFermer 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtOri 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   4035
         TabIndex        =   2
         Text            =   "0"
         Top             =   510
         Width           =   1095
      End
      Begin VB.TextBox txtOri 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtOri 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   720
         TabIndex        =   0
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblDistance 
         Caption         =   "D ="
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblUnite 
         Caption         =   "UnitéAngle"
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblDirection 
         Caption         =   "Direction"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblY 
         Caption         =   "Y ="
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblX 
         Caption         =   "X ="
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmTraject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmTraject   -   TRAJECT.FRM
'   Feuille principale de GIRATION :  Saisie et affichage de la trajectoire

'**************************************************************************************
  
  '*************              Déclarations            *************
Option Explicit


Private contextuel As Boolean
Private AgirSurX As Boolean
Private AgirSurY As Boolean
Private AngleBase As Single
Private AngleBase2 As Single
'Private gtAnglePerp As Boolean

Private OutilZoom As TOOLZOOMEnum
Private SuperZoom As Boolean

Const ID_DEBUT = "DEBUT"
Const ID_INFINI = "INFINI"
Const ID_LIGNE = "LIGNE"
Const ID_RACCORD = "RACCORD"
Const ID_CERCLE = "CERCLE"
Const ID_VIRAGE = "VIRAGE"

Const ID_AngleTotal = "Angle total" ' label totalisation des angles
Const ID_Origine = "Origine"        ' Ecrire au format texte
Const ID_NomsFichiers = "Noms des fichiers" ' Ecrire au format texte

Const IDm_OrigineEnDehors = "L'origine est en-dehors de la fenêtre" & vbCrLf & "Confirmer la modification"
Const ID_OrigineTraject = "Origine de la trajectoire"
Const ID_AngleTracteur = "Angle du tracteur "

Const ID_EnregistrerTrajectoire = "Sauver la " & ID_Trajectoire
Const IDm_TrajectoireIncomplet = "Trajectoire incomplète" & vbCrLf & "Enregistrement partiel jusqu'à la ligne"
Const ID_Rechercher = "Le rechercher"

Const ID_Minimum = "minimum"
Const ID_Maximum = "maximum"
Const ID_Admissible = "admissible"
Const ID_Absolu = "absolu"
Const ID_Strict = "strictement"
Const ID_Superieur = "supérieur à"
Const ID_Inferieur = "inférieur à"

Const ID_ValeurRayon = "Valeur de rayon"

Const IDm_ReinitLigne = "Réinitialiser la ligne"
Const IDm_PrecedentIncomplet = "Ligne précédente incomplète"
Const IDm_PrecedentIncompatible = "Incompatible avec le déplacement précédent"
Const IDm_SuivantIncompatible = "Incompatible avec le déplacement suivant"
Const IDm_ContinuiteRayonImpossible = "La continuité des rayons ne peut être assurée"
Const IDm_RayonFinalProcheDepart = "Le rayon final serait trop voisin du rayon de départ"
Const IDm_2LignesConsecutives = "Pas 2 déplacements LIGNE consécutifs "
Const IDm_2CerclesConsecutifs = "Pas 2 déplacements CERCLE consécutifs"
Const IDm_2RayonsInfinisConsecutifs = "L'angle de giration ne peut être nul" & vbCrLf & "après un déplacement de rayon infini"
Const IDm_ConstructionImpossible = "Construction impossible"
Const IDm_ConstructionImpossibleTouteVitBraq = IDm_ConstructionImpossible & ", quelle que soit la vitesse de braquage"
Const IDm_ErrConstr = "Erreur de construction"
Const IDm_AngleMaxi = "Saisir un angle inférieur à "
Const IDm_LgDepltImportant = "Longueur de déplacement importante"
Const IDm_VitBraqTropGrand = "Vitesse de braquage trop grande"
Const IDm_VitBraqTropFaibleInduite = "Cette valeur conduit à une Vitesse de braquage trop faible"
Const IDm_IncompatibleRayonSigneAngle = "Le signe de l'angle de giration est incompatible avec le rayon"
Const IDm_RayonCompris = "Le rayon doit être compris entre "    ' blanc final important
Const IDm_AngleCompris = "L'angle de giration doit être compris entre " ' blanc final important
Const IDm_MemeSigne = "doivent être de même signe"
Const IDm_SigneContraire = "doivent être de signe contraire"

Const IDm_Augmenter = "Augmenter"
Const IDm_Reduire = "Réduire"
Const ID_AmplitudeRayon = " l'amplitude du rayon" ' blanc initial important
Const ID_AmplitudeAngle = " l'amplitude de l'angle" ' blanc initial important
Const IDm_SansDepasser = "sans dépasser"

Const ID_LaLongueur = "la " & ID_Longueur
Const ID_LaVitBraq = "la " & ID_VitBraq
Const ID_LAngle = "l'angle de giration"
Const ID_PositifLongueur = "positive"
Const ID_NegatifLongueur = "négative"
Const ID_PositifVitBraq = "positive"
Const ID_NegatifVitBraq = "négative"

Const IDm_RedefinirAutreValeur = "Redefinir l'autre valeur"
Const IDm_AugmenterLongueur = IDm_Augmenter & " " & ID_LaLongueur
Const IDm_AugmenterVitBraq = IDm_Augmenter & " " & ID_LaVitBraq
Const IDm_ChangerAngleDeSigne = "Changer le signe de l'angle de giration"
Const IDm_DepltIndetermine = "Déplacement indéterminé" & vbCrLf & "Utiliser les autres paramètres (" & ID_Longueur & ID_Ou & ID_VitBraq & ")"
Const IDm_PrendreRayonPlusProche = "Prendre une valeur plus proche du rayon précédent"
Const ID_EcartRayonPrecedent = " l'écart avec le rayon précédent "  ' Blancs importants
Const IDm_OuLeChangerDeSigne = "ou le changer de signe"
Const IDm_OuChangerDeSigne = "ou changer de signe"
Const IDm_ChangerSigneRayon = "Changer le signe du rayon"
Const IDm_ChangerSigneVitbraq = "Changer le signe de la vitesse de braquage"
Const IDm_ChangerLeSigneDe = "Changer le signe de "   ' blanc final important
Const IDm_AngleNonNul = "Angle de giration non nul"
Const IDm_SupprDebutInterdit = "Pas de suppression de la ligne " & ID_DEBUT
Const IDm_SupprFinInterdit = "Pas de suppression de la dernière ligne"
Const IDm_SupprIrreversible = "La suppression est irréversible"
Const IDm_DecomposIrreversible = "La décomposition de virage est irréversible"
Const IDm_InsertionAudebutInterdit = "Pas d'insertion avant le début"
Const IDm_LimiteAtteinte = "Limite atteinte"
Const IDm_SelectVirage = "Sélectionner un VIRAGE"
Const IDm_SelectLigne = "Sélectionner une ligne"
Const IDm_SelectUneLigne = "Sélectionner une seule ligne"
Const IDm_VirageIncomplet = "Virage incomplet"
Const IDm_ConfirmDecompos = "Confirmer la décomposition"

Const ID_FinProjet = "Fin du projet"  ' Pour version démo

Dim bleupale As Long
Const LIGNE = 0, RACCORD = 1, CERCLE = 2, VIRAGE = 3, DEBUT = 4
Const RAY = 1, ANGGIR = 2, LG = 3, VBRAQ = 4
Const GAUCHE = 0, DROITE = 1

Const FmtGrille = "#0.0##"
Const nbIteration = 500
Const aVehMin = 0.0005
Const MAXLIGNEDEMO = 5
' Un rayon > 10000 est considéré comme infini
Const INFINI = 10000
Private libCol(4) As String

Public flagRecalcul As Boolean
Public nouveau As Boolean, IgnoreOrigine As Boolean
Public flagOriginePosition As Boolean, flagOrigineDirection As Boolean, flagOrigineGuidée As Boolean, flagOrigineLibre As Boolean
Public ExportOK As Boolean
Public Modif As Boolean

'Public TypeZoom As Integer

' Composants d'une ligne de trajectoire
Private Type ligneTraj
  deplacement As Integer
  Rayon As Variant
  DeltaBeta As Variant
  Longueur As Variant
  Vitbraq As Variant
  Alpha As Double           ' angle des roues
  Beta As Double             ' angle du tracteur
  Sigma As Double         ' angle de la 1è remorque
  Gamma As Double       ' angle de la 2è remorque
  p As New PT
  dessinCalculé As Boolean
  dessinEffacé As Boolean
  DernierChampSaisi(1) As Integer
  Modif As Boolean
  Complet As Boolean
  MemoPt() As New PT
  MemoBeta() As Double
  MemoPtG() As New PT
  MemoPtD() As New PT
  numVirage As Integer
End Type

Private TabLigneTraj() As ligneTraj
Private nbVirage As Integer
'Private mesLignesTraj As lTrajs

Private Type ligneVirage
  element(2) As ligneTraj
End Type

Private TabVirage() As ligneVirage

' Mémorisation des positions en vue du pas à pas
Private tabTrajPosition() As Integer, tabPtPosCarross() As Integer
Public maxPosition As Integer
Public VitPas As Integer
Public PasManuel As Boolean
Dim Msg As String
Dim numEnreg As Integer
'Private LgTot As Single, AngleTot As Single

Dim nbdepl%
Dim ligneSel As Integer   ' ligne sélectionnée par l'utilisateur dans la Grille

Public FDP As String, Export As String ' , NomFich As String : Supprimé v3.2 : AV 27/11/2000
'Public colPlans As New Collection
Public colPlans As New Plans
Public GraphFDP As New GRAPHIQUE
Public Carross As New GRAPHIQUE
Public GraphExport As New GRAPHIQUE
Public nbRemorque As Integer
Private wTracteur As INSERT

Private maxDiagonal(2) As Double    ' Diagonale maximale d'un élément de l'attelage

' paramètres d'échelle du dessin
Public POrigine As New PT
Public neworx As Single, newory As Single, diagonalEcran As Single, rayonInscrit As Single
Public Milieux As New Pts, Echelles As New Pts

' Variables de détection et de mémorisation de points cliqués pour les fonctions Zoom et Saisie de l'origine
Private xPremierPt As Single, yPremierPt As Single, flagDeuxiemePt As Boolean, flagTroisiemePt As Boolean, glisser As Boolean
Private xAvPremierPt As Single, yAvPremierPt As Single
' Limites du dessin (coordonnées réelles)
Public pmin As New PT, pmax As New PT, pminFDP As New PT, pmaxFDP As New PT

Public pas_roue As Single, ds As Single, Roues As Boolean

' origine du 1er déplacement
' Indique si l'origine a été déplacée
Private modifOrigine As Boolean
' 1 valeur par élément de carrosserie (TRACTEUR, REMORQUE1, REMORQUE2)
Private sauvOri(2) As String, iniOri(2) As String, tmpOri(2) As String

' véhicule
Public Vehic As New VEHICULE
Public v As Double  ' vitesse
Private e(2) As Double, voav As Double, voar(2) As Double, alphaVehMax As Double, RVehMin As Double, aVehMax As Double, Rmax As Double
Private PAFav(2) As Double, PAFar(2), larg(2) As Double, Dpat(1) As Double
Private mDebord(1) As Double, mSurLarg(2) As Double
Private transRemorque(1) As Double
Private K0Remorque As Double, K1Remorque As Double    ' coefficients pour CalcAngleRemorqueOptimise

' déplacement
Public R As Double, l As Double
Private Alpha As Double, Beta As Double, Alphadeb As Double, betadeb As Double, DeltaBeta As Double, a As Double, coefBeta As Double, Sigma As Double, Gamma As Double

'flags
Public flagZoom As Boolean, flagPAN As Boolean
Public numFichProtégé As Integer


Private Sub cmdCancel_Click()
Dim i As Integer

  If fraOrigine.Enabled Then
    modifOrigine = False
    BasculeGrilleOrigine activeGrille:=True
  
  Else
    MDIGiration.staMesure.Panels("Distance") = ""
    If gtDistance Or gtAngleDyn Or gtOutilZoom = TOOL_PAN Then Exit Sub
    If gtOutilZoom = TOOL_ZOOM Then Form_KeyDown vbKeyEscape, 0: Exit Sub
    
    ExecuteDernierClick 0, 0, InterruptOrigine:=True

    IgnoreOrigine = True
    For i = 0 To 2
      txtOri(i) = tmpOri(i)
    Next
    IgnoreOrigine = False
    DesCarrossPourOrigine
  End If
  
End Sub

'*************       Procédures               *************

Private Sub cmdFermer_Click()
Dim i As Integer, Valeur(1) As Double
Dim newBeta As Single

  
  ' L'utilisateur a fait RETOUR-CHARIOT sur un champ de saisie
  If ActiveControl.name = "txtOri" Then SendKeys "{TAB}": Exit Sub
  
  If Not modifOrigine Then
    BasculeGrilleOrigine True
    Exit Sub
  End If
    
    
  Valeur(0) = trEchelX(CDbl(txtOri(0)), False)
  Valeur(1) = trEchelY(CDbl(txtOri(1)), False)
  If Valeur(0) < 0 Or Valeur(0) > xMaxPicture Or _
      Valeur(1) < 0 Or Valeur(1) > yMaxPicture Then
    Select Case MsgBox(IDm_OrigineEnDehors, vbYesNoCancel + vbQuestion + vbDefaultButton3, ID_OrigineTraject)
    Case vbNo
      colPlans(GirPlanReserve(7)).couleur = picTrajectoire.BackColor   ' Pour ne pas dessiner le marquage du sens
      For i = 0 To 2
        txtOri(i) = iniOri(i)
      Next
      modifOrigine = False
    Case vbCancel
      Exit Sub
    Case vbYes
    End Select
  End If

  BasculeGrilleOrigine True
  
  If modifOrigine Then
    modifOrigine = False
    
    With TabLigneTraj(0)
    'With mesLignesTraj(1)
      newBeta = angConv(wTracteur.rot, radian)
      .Beta = newBeta
      .Sigma = newBeta
      .Gamma = newBeta
      .p.X = wTracteur.p.X
      .p.Y = wTracteur.p.Y
      .MemoPt(0, 0).X = .p.X
      .MemoPt(0, 0).Y = .p.Y
      For i = 0 To nbRemorque
        .MemoBeta(i, 0) = .Beta
        If i > 0 Then
          .MemoPt(i, 0).X = .MemoPt(i - 1, 0).X + transRemorque(i - 1) * Cos(.Beta)
          .MemoPt(i, 0).Y = .MemoPt(i - 1, 0).Y + transRemorque(i - 1) * Sin(.Beta)
        End If
      Next
    End With
      
    dessiner "TOUT", "Origine"
  End If
  
End Sub

Public Sub BasculeGrilleOrigine(ByVal activeGrille As Boolean)
Dim i As Integer

  If Not modifOrigine And activeGrille Then
    For i = 0 To nbRemorque
      FinDesCarross TabLigneTraj(0), 0, i
    Next
  End If

  Grille.Visible = activeGrille
  fraOrigine.Visible = Not activeGrille
  ' Modif AV v3.2.1 : 04/04/2000 : les boutons de Zoom sont en barre d'outils, et c'est mieux d'inactiver que d'invisibiliser (uniqt bouton Origine)
'  picBoutons.Enabled = activeGrille
  
  With MDIGiration
      'Fichier Imprimer
    .mnuFich(MNUIMPRIMER).Enabled = activeGrille And Printers.Count > 0
    .tbrGiration.Buttons("btnPrint").Enabled = .mnuFich(MNUIMPRIMER)
  End With
  
  If activeGrille Then
    colPlans(GirPlanReserve(7)).couleur = picTrajectoire.BackColor
    HelpContextID = GENERALITESTRAJ
    For i = 1 To (nbRemorque + 1) * 4 + 2
      linCarrosserie(i).Visible = False
    Next
    For i = 0 To nbRemorque
      shpPcarac(i).Visible = False
    Next
    fraAngleTotal.Visible = MDIGiration.mnuOutils(MNUANGLETOTAL).Checked
    
  Else
    txtOri(0).SetFocus
    colPlans(GirPlanReserve(7)).couleur = OptionAffichage(NumFille).Carross(TRACTEUR).couleur
    HelpContextID = SAISORI
    For i = 1 To (nbRemorque + 1) * 4 + 2
      linCarrosserie(i).Visible = True
      If i <= (nbRemorque + 1) * 4 Then
        linCarrosserie(i).BorderColor = OptionAffichage(NumFille).Carross((i - 1) \ 4).couleur
      Else
        linCarrosserie(i).BorderColor = OptionAffichage(NumFille).Carross(TRACTEUR).couleur
      End If
    Next
    For i = 0 To nbRemorque
      shpPcarac(i).Visible = True
      shpPcarac(i).BorderColor = OptionAffichage(NumFille).Carross(i).couleur
    Next
    linMarque(0).Visible = False
    linMarque(1).Visible = False
    fraAngleTotal.Visible = False
  End If

  MDIGiration.HelpContextID = HelpContextID
  
End Sub

Private Sub Form_Deactivate()
  OutilZoom = gtOutilZoom
End Sub


Private Sub Form_Initialize()
  OutilZoom = TOOL_SANSZOOM
End Sub

Private Sub Grille_KeyPress(KeyAscii As Integer)
  If KeyAscii > 47 And KeyAscii < 127 Then
    AfficheAngleTot Grille.ActiveRow - 1
    lblAngleTotal.Enabled = False
  End If
End Sub

Private Sub picTrajectoire_LostFocus()
  MDIGiration.staBarre.Panels("X").Text = ""
  MDIGiration.staBarre.Panels("Y").Text = ""
End Sub

Private Sub txtOri_Change(Index As Integer)
  ' Changement de l'origine : redessiner le tracteur
Dim Valeur As Double

  If Not (IsNumeric(txtOri(0)) And IsNumeric(txtOri(1)) And IsNumeric(txtOri(2))) Then Exit Sub
    
    ' En saisie interactive de l'origine pas de dessin du tracteur
      If flagOrigineDirection Then
        If nbdepl > 1 Then
          nbdepl = 0
        Else
          Exit Sub
        End If
      End If
    
    modifOrigine = True
    
    If Index = 0 Then
      Valeur = trEchelX(CDbl(txtOri(0)), False)
      AgirSurX = False
      If Valeur < 0 Or Valeur > xMaxPicture Then Exit Sub
    ElseIf Index = 1 Then
      Valeur = trEchelY(CDbl(txtOri(1)), False)
      AgirSurY = False
      If Valeur < 0 Or Valeur > yMaxPicture Then Exit Sub
    End If
        
    
  If Carross.gphInserts.Count > 0 Then
        ' Mémorisation de la nouvelle
    Select Case Index
    Case 0
'      Carross.Inserts.Item(1).p.X = CDbl(txtOri(0))
      wTracteur.p.X = CDbl(txtOri(0))
    Case 1
'      Carross.Inserts.Item(1).p.Y = CDbl(txtOri(1))
      wTracteur.p.Y = CDbl(txtOri(1))
    Case 2
'      Carross.Inserts.Item(1).rot = CDbl(txtOri(2))
      wTracteur.rot = CDbl(txtOri(2))
    End Select
  End If

If Index = 0 And AgirSurY Then Exit Sub
          ' Dessin de la nouvelle
'  If Not IgnoreOrigine Then DesCarross TRACTEUR      ' v 3.0
    If Not IgnoreOrigine Then DesCarrossPourOrigine   '<-- v anglaise (3.1)
  'If Not IgnoreOrigine And (Index <> 0 Or AgirSurX) Then DesCarrossPourOrigine: If Index <> 2 Then DoEvents

  End Sub

'Suite à un Zoom, on a perdu les coordonnées de l'origine qui est en cours de modif
Public Sub RedefOrigine()
  wTracteur.p.X = CDbl(txtOri(0))
  wTracteur.p.Y = CDbl(txtOri(1))
  wTracteur.rot = CDbl(txtOri(2))
  DesCarrossPourOrigine

End Sub

Private Sub cmdBoutonOrigine_Click()
Dim i%

'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect
  
  If fraOrigine.Visible Then
    cmdCancel = True
  Else
  
    BasculeGrilleOrigine False
    
    ' Pour repositionnement sur l'élément DEBUT et Effacement de l'ancienne carrosserie
    gtEffacement = True
    For i = 0 To nbRemorque
      FinDesCarross TabLigneTraj(0), 0, i
    Next
    gtEffacement = False
    
    DesCarrossPourOrigine
    
    For i = 0 To 2
      iniOri(i) = txtOri(i)
    Next
    
  End If
  
End Sub

Private Sub txtOri_GotFocus(Index As Integer)
  
  sauvOri(Index) = txtOri(Index)
  txtOri(Index).SelStart = 0
  txtOri(Index).SelLength = Len(txtOri(Index))
  
End Sub

Private Sub txtOri_LostFocus(Index As Integer)
Dim AngleMax As String

      If Not IsNumeric(txtOri(Index)) Then
        MessA IDm_Numerique
        txtOri(Index) = sauvOri(Index)
        txtOri(Index).SetFocus
' Vérification des valeurs saisies pour éviter les plantages en réouverture
' d'épure si valeur trop grande pour l'origine
  Else
    Select Case Index
    Case 0, 1   ' X,Y
      On Error GoTo GestErr
      If Abs(CCur(txtOri(Index))) >= 10000000 Then
        MessA IDm_StrictCompris & " -10000000 " & ID_Et & " 10000000"
        txtOri(Index) = sauvOri(Index)
        txtOri(Index).SetFocus
      End If
    
    Case 2    ' Angle
        AngleMax = CStr(eqvPI(OptGen.unite))
        If Abs(CCur(txtOri(Index))) > eqvPI(OptGen.unite) Then
          MessA IDm_Compris & "-" & AngleMax & ID_Et & AngleMax
          txtOri(Index) = sauvOri(Index)
          txtOri(Index).SetFocus
      End If
    End Select
    
    ' MessA "Valeur origine :" & txtOri(Index)
    ' CERTU/ESi/GNSI le 10/01/2014 Ajout de la conversion de la coordonnee en Double
        
    txtOri(Index) = Arrond(CDbl(txtOri(Index)), 2)
  
  End If
  
  Exit Sub
  
GestErr:
  If Err = 6 Then
    txtOri(Index) = Sgn(txtOri(Index)) * 10000000
    Resume 0
  End If
End Sub

Private Sub picTrajectoire_DblClick()
  
If Not MDIGiration.Enabled Then Exit Sub
If gtInteractionEnCours Then Exit Sub

  frmGrandDessin.Show vbModal
  
End Sub

'**************************************************************************
' Déplacement de la souris sur le dessin : Affichage des coordonnées
' Si Action particulière (Zom, Pan, Distance...) : Curseur élastique
'**************************************************************************
Private Sub picTrajectoire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static X1 As Single, Y1 As Single
  Dim p1 As New PT, p2 As New PT, p3 As New PT

  If Not fCourante Is Me Then Exit Sub
  
  If gtFacteurZoom <> 0 And gtCoordActif Then
    MDIGiration.staBarre.Panels("X").Text = "X = " & Format(trEchelX(X, toReel:=True), "#0.0#")
    MDIGiration.staBarre.Panels("Y").Text = "Y = " & Format(trEchelY(Y, toReel:=True), "#0.0#")
  End If

  If flagDeuxiemePt Then       ' frmCadrage chargé et 1 pt déjà cliqué ou saisie de l'origine et pt origine déjà cliqué
    X1 = X
    Y1 = Y
    If flagZoom Then
        shpZoom.Left = Min(xPremierPt, X)
        shpZoom.Top = Min(yPremierPt, Y)
        shpZoom.Width = Abs(xPremierPt - X)
        shpZoom.Height = Abs(yPremierPt - Y)
    Else
        linPan.X2 = X
        linPan.Y2 = Y
    End If
    glisser = (Button And vbLeftButton) > 0
    nbdepl = nbdepl + 1
  End If
    
  If X = xPremierPt And Y = yPremierPt Then Exit Sub  'Move déclenché juste après le Up
    
  Dim wAngle As Single, wDistance As Single
  If flagOriginePosition Then
    ' Saisie guidée : recalcul du 2è pt comme projection du point courant sur la droite déterminée par les 2 premiers (AngleBase)
    If flagOrigineGuidée And flagTroisiemePt Then
      p1.X = xPremierPt
      p1.Y = yPremierPt
      p2.X = X
      p2.Y = Y
      wAngle = CalcAngle(p1, p2) - AngleBase
      wDistance = Distance(xPremierPt, X, yPremierPt, Y) * Cos(wAngle)
      p3.X = wDistance
      p3.Y = 0
      Set p2 = RotTrans(p3, p1, AngleBase)
      linPan.X1 = p2.X
      linPan.Y1 = p2.Y
      linAngle.X2 = p2.X
      linAngle.Y2 = p2.Y
      wDistance = Distance(xPremierPt, X, yPremierPt, Y) * Cos(wAngle - pi / 2)
      txtDistance = Format(Abs(wDistance) / gtFacteurZoom, "#0.#")
    End If
    AgirSurX = (txtOri(0) <> Arrond(trEchelX(X, toReel:=True), 2))
    AgirSurY = (txtOri(1) <> Arrond(trEchelY(Y, toReel:=True), 2))
    If AgirSurX Then txtOri(0) = Arrond(trEchelX(X, True), 2): DoEvents
    If AgirSurY Then txtOri(1) = Arrond(trEchelY(Y, True), 2)
    DoEvents
    'If AgirSurX Or AgirSurY Then DoEvents
  End If
    
  If flagOrigineDirection Or ((gtDistance Or gtAngleDyn) And flagDeuxiemePt) Then
      p1.X = xPremierPt
      p1.Y = -yPremierPt
      p2.X = X
      p2.Y = -Y
      wAngle = CalcAngle(p1, p2)
      If flagOrigineDirection Then  ' Affichage de la direction
        txtOri(2) = Arrond(angConv(wAngle, unite), 2)
      ElseIf Not flagTroisiemePt Then ' Attente du 2è pt si Distance ou Angle
        MDIGiration.staMesure.Panels("Distance").Text = "D = " & Format(Distance(X, xPremierPt, Y, yPremierPt) / gtFacteurZoom, "#0.0#") & _
          " - A = " & Format(angConv(wAngle, unite), "#0.0#")
      ElseIf gtAngleDyn Then  ' Attente du 3è pt si Angle
        p3.X = xAvPremierPt
        p3.Y = -yAvPremierPt
        wAngle = wAngle - CalcAngle(p1, p3)
        ' wAngle appartient à ]-pi,+pi]
        If wAngle < -pi Then wAngle = wAngle + 2 * pi
        MDIGiration.staMesure.Panels("Distance").Text = "Angle = " & _
        Format(angConv(wAngle, unite), "#0.0#") & " " & libUnite(OptGen.unite)
      End If
  End If
 
End Sub

'**************************************************************************
' Clic bouton droit sur la Trajectoire : Menu contextuel Outils
' Clic bouton gauche : Zoom - PAN - Modif graphique de l'origine
'**************************************************************************

Private Sub picTrajectoire_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

  If Button = vbLeftButton And gtFacteurZoom <> 0 And Not gtCoordActif And gtOutil <> TOOL_AUCUN Then
    MDIGiration.staBarre.Panels("X").Text = "X = " & Format(trEchelX(X, toReel:=True), "#0.0#")
    MDIGiration.staBarre.Panels("Y").Text = "Y = " & Format(trEchelY(Y, toReel:=True), "#0.0#")
  End If
    
  If flagZoom And Button = vbRightButton Then
    If Not flagDeuxiemePt Then PopupMenu MDIGiration.mnuZoomMul: Exit Sub
  End If
  
  If flagZoom Or flagPAN Or flagOrigineLibre Or flagOrigineGuidée Then
    If Button = vbLeftButton Then
      ' frmCADRAGE chargée ou Saisie de l'origine en cours
      If flagDeuxiemePt Then
        ExecuteDeuxiemeClick X, Y
      Else                       ' 1er pt
        nbdepl = 0
        flagDeuxiemePt = True
        xPremierPt = X
        yPremierPt = Y
        If flagZoom Then
          With shpZoom
            .Visible = True
            .Left = X
            .Top = Y
            .Width = 0
            .Height = 0
          End With
          bgTexte ""
          bgTexte ID_Deuxieme & " " & ID_Coin & " : "
          
        Else
          With linPan
            .Visible = True
            .X1 = X
            .Y1 = Y
            .X2 = X
            .Y2 = Y
          End With
          If flagOrigineGuidée Then  ' Saisie de l'origine par 3 points (l'angle en 1er)
            flagOrigineDirection = True
            bgTexte ""
            bgTexte ID_AngleTracteur & " : "
          ElseIf flagOriginePosition Then ' Saisie de l'origine par 2 points (la position en 1er)
            flagOriginePosition = False
            txtOri_Change 0
            txtOri_Change 1
            flagOrigineDirection = True
            bgTexte ""
            bgTexte ID_AngleTracteur & " : "
          ElseIf flagPAN Then ' Couvre en principe tous les autres cas que saisie de l'origine
            gtDistance = (gtOutil = TOOL_DIST And gtOutilZoom = TOOL_SANSZOOM)
            gtAngleDyn = (gtOutil = TOOL_ANGLEDYN And gtOutilZoom = TOOL_SANSZOOM)
            If fraOrigine.Visible And (gtDistance Or gtAngleDyn) Then fraOrigine.Enabled = False
            bgTexte ""
            bgTexte ID_Deuxieme & " " & ID_Point & " : "
          End If
        End If
  
      End If
    End If
  
  ' Bouton droite : menu contextuel
  ElseIf Button = vbRightButton And Grille.Visible Then
    With MDIGiration
      For i = 1 To 4
        MDIGiration.mnuOutils(i).Visible = False
      Next
      PopupMenu MDIGiration.mnuBarre(1), , , , .mnuOutils(MNURAFRAICHIR)
      For i = 1 To 4
        MDIGiration.mnuOutils(i).Visible = True
      Next
    End With
  End If
  
End Sub
  
  '**************************************************************************
  ' Clic bouton droit sur la Trajectoire : Ignoré
  ' Clic bouton gauche : Zoom - PAN - Modif graphique de l'origine
  '**************************************************************************
  Private Sub picTrajectoire_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim p1 As New PT, p2 As New PT
    
  If flagDeuxiemePt And glisser And X <> xPremierPt And Y <> yPremierPt And nbdepl > 10 Then
    If Button = vbLeftButton Then ExecuteDeuxiemeClick X, Y
  End If
      
End Sub

'**************************************************************************************************
' Exécution  du deuxième Clic souris dans une commande interactive (Zoom,Pan,Mesure, Saisie Origine)
' Cette procédure est également appelée lors du 3è clic (flagTroisiemePt=True) pour les commandes qui le nécessitent
'**************************************************************************************************

Private Sub ExecuteDeuxiemeClick(ByVal X As Single, ByVal Y As Single)
Dim ExecTroisiemePt As Boolean

  ExecTroisiemePt = (gtAngleDyn Or flagOrigineGuidée) And Not flagTroisiemePt
  
    If ExecTroisiemePt Then ' Préparation de la saisie d'un 3è pt : Origine guidée ou Angledyn
        flagTroisiemePt = True  ' l'objet linAngle devient visible et prend la position de linPAN, qui va continuer à bouger
        linAngle.Visible = True
        linAngle.X1 = linPan.X1
        linAngle.Y1 = linPan.Y1
        linAngle.X2 = linPan.X2
        linAngle.Y2 = linPan.Y2
        xAvPremierPt = xPremierPt
        yAvPremierPt = yPremierPt
        xPremierPt = X
        yPremierPt = Y
        linPan.X1 = X
        linPan.Y1 = Y
        If flagOrigineGuidée Then
          txtDistance.Visible = True
          lblDistance.Visible = True
          flagOrigineDirection = False
    '      txtOri_Change 2
          flagOriginePosition = True
          bgTexte ID_OrigineTraject ' "Origine du tracteur : "
          Dim p1 As New PT, p2 As New PT
          p1.X = xAvPremierPt
          p1.Y = yAvPremierPt
          p2.X = X
          p2.Y = Y
          AngleBase = CalcAngle(p1, p2)   ' +pi/2
        Else  ' AngleDyn
          bgTexte ""
          bgTexte ID_Dernier & " " & ID_Point & " : "
        End If
    
    Else
      flagDeuxiemePt = False
      glisser = False
      If flagZoom Then
        flagZoom = False
        shpZoom.Visible = False
        GirStand.ExecuteZoom Min(X, xPremierPt), Max(Y, yPremierPt), Max(X, xPremierPt), Min(Y, yPremierPt)
      Else
        ExecuteDernierClick X, Y
      End If
      xPremierPt = 0
      yPremierPt = 0
    End If

End Sub

Private Sub ExecuteDernierClick(ByVal X As Single, ByVal Y As Single, Optional ByVal InterruptOrigine As Boolean)
Dim i As Integer

    If InterruptOrigine Then
      flagDeuxiemePt = False
    Else
      If flagOrigineGuidée Or flagOrigineDirection Then
        For i = 0 To 2
          tmpOri(i) = txtOri(i)
        Next
      End If
    End If
   
    If gtOutilZoom = TOOL_PAN Or gtDistance Or flagTroisiemePt Then
           'PAN, gtDistance, ou 3è pt (AngleDyn ou Saisie Ori 3pts)
      flagTroisiemePt = False
      linAngle.Visible = False
      flagPAN = False
      linPan.Visible = False
      If gtOutilZoom = TOOL_SANSZOOM And flagOrigineGuidée Then ' saisie Origine par 3 pts
        flagOriginePosition = False
        flagOrigineGuidée = False
        txtDistance.Visible = False
        txtDistance = ""
        lblDistance.Visible = False
        bgTexte ""
        fraOrigine.Enabled = True
        DefautCurseur
      Else  'PAN, gtDistance ou gtAngleDyn
        GirStand.ExecuteZoom xPremierPt, yPremierPt, X, Y
      End If
    
    Else    ' Saisie Origine par 2 points
      bgTexte ""
      fraOrigine.Enabled = True
      linPan.Visible = False
      flagOrigineDirection = False
      flagOrigineLibre = False
      If InterruptOrigine Then
        flagOriginePosition = False
      Else
        txtOri_Change 2
      End If
      DefautCurseur
    End If
    
    
End Sub

Private Sub cmdSaisieOri_Click(Index As Integer)
Dim i As Integer

  Select Case Index
  Case 0
    flagOrigineLibre = True
    flagOriginePosition = True
    bgTexte ID_OrigineTraject ' "Origine du tracteur : "
  Case Else
    flagOrigineGuidée = True
    bgTexte ID_Premier & " " & ID_Point & " : "
  End Select
  
  MDIGiration.staMesure.Panels("Distance") = ""
  For i = 0 To 2
    tmpOri(i) = txtOri(i)
  Next
  ObjetDessin.MousePointer = vbCrosshair ' Croix
  fraOrigine.Enabled = False
End Sub

Private Sub cmdVéhicule_Click()
    
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect
Dim fVéhicule As New frmVéhicule

  fVéhicule.Show vbModal, MDIGiration
  Set fVéhicule = Nothing
  
End Sub

'***************************************************************************
'           Procédures relatives à la feuille (ouverture,activation,fermeture)
'****************************************************************************

Private Sub Form_Load()
' cette procédure est appelée depuis frmTrajpar (Nouveau)
Dim i As Integer, h As Single, htot As Single
Dim numRemorque As Integer
Dim Nombase As String
Dim p As New PT

DégraisserFonte Me

If FichierJournal Then
  libCol(0) = "Rayon"
  libCol(1) = "Angle"
  libCol(2) = "Longueur"
  libCol(3) = "Vit. Braq."
End If

bleupale = RGB(0, 255, 255)

SetDeviceIndependentWindow Me

lblUnite = libUnite(OptGen.unite)

'Initialisation de l'aide pour VB5, qui ne considère pas que c'est la feuille MDI qui a le focus  si la grille est active
HelpContextID = GENERALITESTRAJ

  ' paramètres d'origine et d'échelle du dessin
gtOrx = picTrajectoire.ScaleWidth / 2
gtOry = picTrajectoire.ScaleHeight / 2
POrigine.X = gtOrx
POrigine.Y = gtOry
neworx = 0
gtFacteurZoom = 0

' Initialisation de la grille
With Grille
' Définition des libellés de la combo en première colonne
  .Row = -1
  .Col = 1
  .CellType = 8   ' Combo
  .TypeComboBoxList = ID_LIGNE & Chr(9) & ID_RACCORD & Chr(9) & ID_CERCLE & Chr(9) & ID_VIRAGE
  
  .MaxRows = 1
  
' Initialisation de la première ligne : DEBUT
  .Row = 1
  .Col = 1
  .Lock = True
  .CellType = 1 ' Texte
  .Col = 3
  .ForeColor = vbRed
  .Col = 4
  .ForeColor = vbRed
  
  For i = 3 To 5: .Col = i: .Lock = True: .BackColor = grisPale: Next '.BackColor = QBColor(8): Next
  .SetText 1, 1, ID_DEBUT
End With

'Initialisation des options d'affichage de la trajectoire
'avec les valeurs lues dans Giration.ini

Me.Tag = NbreFilles + 1
OptionAffichage(NumFille) = DefOptAff
pas_roue = OptionAffichage(NumFille).pas(0) / 5


If nouveau Then ' Nouvelle trajectoire
  ' caractéristiques du véhicule
  ColVéhicules(UCase(frmTrajpar.lstNomVehic)).Clone Vehic
  
    'lireVehicul
  With Vehic
    .nbUtilisation = .nbUtilisation + 1     ' afin que le véhicule ne soit pas modifié tq la trajectoire l'utilise
    nbRemorque = .nbRemorque
    For numRemorque = 0 To nbRemorque
      PAFav(numRemorque) = .PAFav(numRemorque)
      e(numRemorque) = .e(numRemorque)
      PAFar(numRemorque) = .PAFar(numRemorque)
      larg(numRemorque) = .l(numRemorque) / 2
      voar(numRemorque) = .LEar(numRemorque) / 2
      If numRemorque > 0 Then
        Dpat(numRemorque - 1) = .Dpat(numRemorque - 1)
        transRemorque(numRemorque - 1) = -(e(numRemorque - 1) + Dpat(numRemorque - 1))
      End If
    Next
    voav = .LEav / 2
    RVehMin = .RVehMin
    alphaVehMax = Atn(e(0) / (RVehMin + voar(0)))      ' Référence: Dossier de programmation - Bibliothèque de véhicules - §1
    aVehMax = .aVehMax
  End With
  
    ' Vitesse
  v = frmTrajpar.txtVitesse / 3.6         ' Km/h --> m/s
 
Else  ' Ouvrir Trajectoire
  On Error GoTo GestErr
  If lire Then Err.Raise 500
End If

lblCaracVeh = " " & Trim(Vehic.Nom) & "     " & Arrond(v * 3.6, 2) & " " & ID_UniteVitesse

' pas de calcul de la trajectoire tous les cm (signé)
  ds = Sgn(v) * 0.01


' Tableau des positions pour le pas à pas et la procédure Rafraichir
ReDim tabTrajPosition(0) As Integer
ReDim tabPtPosCarross(0) As Integer
tabTrajPosition(0) = 0
tabPtPosCarross(0) = 0

If nouveau Then    'initialisation du 1er enregt
  ReDim TabLigneTraj(0)
  With TabLigneTraj(0)
    .deplacement = DEBUT
    .Rayon = 0
    .DeltaBeta = 0
    .Longueur = 0
    .Vitbraq = 0
    .Complet = True
    ReDim .MemoPt(nbRemorque, 0)
    ReDim .MemoBeta(nbRemorque, 0)
    For numRemorque = 1 To nbRemorque
      .MemoPt(numRemorque, 0).X = .MemoPt(numRemorque - 1, 0).X + transRemorque(numRemorque - 1)
    Next
    .Complet = True
  End With
  
  creLigne    ' Création du 1er déplacement
  Grille.Row = 1
  Alpha = 0
  R = R_alpha("alpha")
End If

numEnreg = 0
Grille.Row = 1
Grille.Col = 2
Grille.Action = 0      'Activation du champ rayon de l'enregistrment DEBUT

If nouveau Then
  Unload frmTrajpar
  Grille.Visible = False
  iniOri(0) = "0"
  iniOri(1) = "0"
  iniOri(2) = "0"
Else
  ' Modif AV v3.2.1 : 04/04/2000 : les boutons de Zoom sont en barre d'outils et l'inactivation est préférable à l'invisibilité
'  cmdBoutonOrigine.Visible = True
'  cmdVéhicule.Visible = True
'  cmdZoom.Visible = True
'  cmdPAN.Visible = True
'  cmdZAvant.Visible = True
End If

' Le dossier de programmation définit k=1+dpat/E
' On en déduit a1 = (k-1)/Er = dpat/E/Er
' par aillleurs a0=1/Er

If nbRemorque > 0 Then
  ' coefficients utilisés par CalcAngleRemorqueOptimise : Constantes liées au véhicule
  K0Remorque = 1 / e(0)
  K1Remorque = Dpat(0) / e(0) / e(1)
End If

Exit Sub

GestErr:
  If Err = 500 Then
    Me.Tag = ""
  Else
    ErrGeneral
  End If

End Sub     ' Form_Load

Public Sub Form_Activate()
    Dim echel As New PT
    Dim p As MSComctlLib.Panel
    
' On utilise Me.Tag pour identifier la fenêtre trajectoire active
' Me.Tag a un numéro (attribué séquentiellement par Form_Load)
 
    gtOutilZoom = OutilZoom
    
    If FichierJournal Then Write #numFichLog, "Activation Trajectoire " & Me.Tag
    
    NumFille = val(Me.Tag) - 1
    Set fCourante = Me
    Set ObjetDessin = fCourante.picTrajectoire
    Set gtPlans = fCourante.colPlans
    For Each p In MDIGiration.staBarre.Panels
      p.Text = ""
    Next
    MDIGiration.staMesure.Panels("Distance") = ""
        
    If Grille.Visible Then
      If Grille.Enabled Then Grille.SetFocus
    End If
        
    AffecterOptAffichage
    
    GriserMenus Etat:=True

    If fraOrigine.Visible Then
      HelpContextID = SAISORI
    Else
      HelpContextID = GENERALITESTRAJ
    End If
    MDIGiration.HelpContextID = Me.HelpContextID

    If Milieux.Count <> 0 Then
    ' Alimentation des variables (déclarées dans GraphStand)  utiles aux fonctions de mise à l'échelle
      Set gtMil = Milieux(Milieux.Count)
      Set echel = Echelles(Echelles.Count)            ' echel est une variable locale, pour déterminer gtFacteurZoom
      gtFacteurZoom = Min(echel.X, echel.Y)
      xMaxPicture = picTrajectoire.ScaleWidth
      yMaxPicture = picTrajectoire.ScaleHeight
      gtOrx = POrigine.X
      gtOry = POrigine.Y
      
    End If
        
  With MDIGiration
        ' On ne laisse visible dans la barre d'outils que les boutons 'Standard'
    .tbrGiration.Buttons("btnZoom").Visible = True
    .tbrGiration.Buttons("btnZAvant").Visible = True
    .tbrGiration.Buttons("btnZArriere").Visible = True
'    .tbrGiration.Buttons("btnZMultiple").Visible = True
    .tbrGiration.Buttons("btnPAN").Visible = True
    If .mnuOutils(MNUMESURES).Checked Then ' Outils de mesure
      .staBarre.Visible = True
      .tbrGiration.Buttons("btnFleche").Visible = True
      .tbrGiration.Buttons("btnXY").Visible = True
      .tbrGiration.Buttons("btnDistance").Visible = True
      .tbrGiration.Buttons("btnAngle").Visible = True
      'Modif v3.3(25/04/07) : Utilisation d'un énumérateur au lieu de constantes
      'If gtOutil >= 0 And gtOutil < 4 Then
      If gtOutil <> TOOL_AUCUN Then
        .tbrGiration.Buttons(gtBoutons(gtOutil)).Value = tbrPressed
      End If
      .staMesure.Visible = (gtOutil = TOOL_DIST Or gtOutil = TOOL_ANGLEDYN)
  Else
      .staBarre.Visible = False
      .tbrGiration.Buttons("btnFleche").Visible = False
      .tbrGiration.Buttons("btnXY").Visible = False
      .tbrGiration.Buttons("btnDistance").Visible = False
      .tbrGiration.Buttons("btnAngle").Visible = False
      .staMesure.Visible = False
    End If
    fraAngleTotal.Visible = .mnuOutils(MNUANGLETOTAL).Checked And Grille.Visible
  End With
  
    If Not flagDeuxiemePt Then
      DefautCurseur
    End If

End Sub

Public Sub AffecterOptAffichage()
  With OptionAffichage(NumFille)
    Roues = .RouAvt.ADessiner Or .RouArr(0).ADessiner
    If nbRemorque > 0 Then
      Roues = Roues Or .RouArr(1).ADessiner
      If nbRemorque = 2 Then
        Roues = Roues Or .RouArr(2).ADessiner
      End If
    End If
    pas_roue = .pas(0) / 5
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---  Demande de sauvegarde avant déchargement de la feuille
'---- Si Cancel=True, la feuille ne sera pas déchargée

Dim reponse As Integer, flag As Integer


On Error GoTo GestErr

If Not MDIGiration.mnuBarre(0).Enabled Then Cancel = True: Exit Sub

'If ErreurGenerale Or gtProtectRetirée Then flag = vbYesNo + vbQuestion Else flag = vbYesNoCancel + vbQuestion
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
If ErreurGenerale Then flag = vbYesNo + vbQuestion Else flag = vbYesNoCancel + vbQuestion

If Me.Tag = "" Then Exit Sub

Form_Activate

If Modif And Not VersionDemo Then
  reponse = MsgBox(ID_EnregistrerTrajectoire & " " & Me.Caption, flag)
Else
  reponse = vbNo
End If

Select Case reponse
Case vbYes
  If FDP <> "" And dlgFDP.FileName = "" Then
      reponse = MsgBox(IDm_EnregistrerFDP & " " & Me.Caption, flag)
    Select Case reponse
    Case vbYes
      EnregFDPSous Me, Cancel
    Case vbCancel
      Cancel = True
    End Select
  End If
  If Not Cancel Then Enregistrer Me, Cancel

Case vbCancel
  Cancel = True
End Select

  If ErreurGenerale Then Cancel = False

  Exit Sub
  
GestErr:
  
    If Not ErreurGenerale Then
      ErrGeneral
    Else
      End
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer, j As Integer, reponse As Integer
      
If FichierJournal Then Write #numFichLog, "Fermeture Trajectoire " & Me.Tag
        
        ' Restaure l'état initial de protection du véhicule
If nouveau Then Vehic.nbUtilisation = Vehic.nbUtilisation - 1
    
         ' Réinitialisation de l'aide
  MDIGiration.HelpContextID = SOMMAIRE  ' 0

  With MDIGiration
        ' On ne laisse visible dans la barre d'outils que les boutons 'Standard'
    .staBarre.Visible = False
    .staMesure.Visible = False
    .tbrGiration.Buttons("btnZoom").Visible = False
    .tbrGiration.Buttons("btnZAvant").Visible = False
    .tbrGiration.Buttons("btnZArriere").Visible = False
'    .tbrGiration.Buttons("btnZMultiple").Visible = False
    .tbrGiration.Buttons("btnPAN").Visible = False
    .tbrGiration.Buttons("btnFleche").Visible = False
    .tbrGiration.Buttons("btnXY").Visible = False
    .tbrGiration.Buttons("btnDistance").Visible = False
    .tbrGiration.Buttons("btnAngle").Visible = False
  
      ' Libération de la mémoire
    .MousePointer = vbHourglass  ' sablier
  End With
  
On Error Resume Next

' ok

  #If Win32 Then        ' Si aucun virage, l'exécutable en code natif W95 se plante sur le Erase
    On Error GoTo Efface
  #End If

'  For i = 0 To UBound(TabLigneTraj)
'    With TabLigneTraj(i)
'      Erase .MemoPt
'      Erase .MemoBeta
'      Erase .MemoPtG
'      Erase .MemoPtD
'    End With
'  Next
  
'  If FichierJournal Then
'    Write #numFichLog, "erase tablignetraj ok "
'    Close #numFichLog
'    FichierJournal = False
'  End If
  
Efface:
  
'  If FichierJournal Then
'    Write #numFichLog, "erase tablignetraj pas ok " & i
'    Close #numFichLog
'    FichierJournal = False
'  End If
  Erase TabLigneTraj

  #If Win32 Then        ' Si aucun virage, l'exécutable en code natif W95 se plante sur le Erase
    On Error GoTo EffaceFDP
  #End If
    
'  For i = 0 To UBound(TabVirage)
'    With TabVirage(i)
'      For j = 0 To 2
'        With .element(j)
'          Erase .MemoPt
'          Erase .MemoBeta
'          Erase .MemoPtG
'          Erase .MemoPtD
'        End With
'      Next
'    End With
'  Next
  
'    If FichierJournal Then
'      Write #numFichLog, "erase ok"
'      Close #numFichLog
'      FichierJournal = False
'    End If
  
  Erase TabVirage
  
EffaceFDP:

'pas ok
'    If FichierJournal Then
'      Write #numFichLog, "erase pas ok"
'      Close #numFichLog
'      FichierJournal = False
'    End If

  With GraphFDP
    With .gphPoints
      While .Count > 1
        .Remove 1
      Wend
    End With

    With .gphLines
      While .Count > 1
        .Remove 1
      Wend
    End With

    With .gphPolylines
      While .Count > 1
        .Remove 1
      Wend
    End With

    With .gphArcs
      While .Count > 1
        .Remove 1
      Wend
    End With

    With .gphCercles
      While .Count > 1
        .Remove 1
      Wend
    End With

    With .gphInserts
      While .Count > 1
        .Remove 1
      Wend
    End With
    
  End With

  GriserMenus Etat:=False
  Set fCourante = Nothing
  Set ObjetDessin = Nothing
  MDIGiration.MousePointer = vbDefault

  If numFichProtégé <> 0 Then Close numFichProtégé
  If Tag <> "" Then gtFeuillesTraject.Remove Tag
  
End Sub

' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré (indicateur : flagVirgule), on détecte la frappe du point décimal
' mais seule la fonction KeyPress semble en mesure de réafficher la virgule
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyEscape And gtOutilZoom = TOOL_ZOOM And Not flagDeuxiemePt Then
    bgTexte ""
    Unload frmCadrage
    flagZoom = False
    flagPAN = False
    gtOutilZoom = TOOL_SANSZOOM ' Ceci permet à DefautCurseur de réinitialiser l'outil de mesure
    DefautCurseur
  End If
  
  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' L'utilisateur vient de frapper un point décimal en configuration virgule pour celui-ci sous Windows
' on écrit une virgule à la place
  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False
End Sub

Private Sub Form_Resize()
Dim savescale As Variant, B_Top As Variant, B_Left As Variant

If Me.Tag = "" Then Exit Sub
  
If FichierJournal Then Write #numFichLog, "Resize Trajectoire ", val(Me.Tag)
'If NumFille <> val(Me.Tag) - 1 Then Exit Sub
If WindowState = vbMinimized Then Exit Sub

'''Ne rien faire si icone
''If ScaleHeight < 2 * Grille.Height Then Exit Sub
''If ScaleWidth < 1000 Then Exit Sub
''
'''Comme ScaleWidth est en Twips, et que Left est en
'''unité Locales : on prend les unités avec ScaleMode
''savescale = ScaleMode
''ScaleMode = vbTwips

'Comme la Grille est toujours en bas, son Top est :
B_Top = ScaleHeight - Grille.Height * (1.05) ' : modif AV 7/11/97
'decalTop = Grille.Top - B_Top
'La position Left n'est jamais modifiée
B_Left = Grille.Left
'Move Gauche, Haut , Largeur, Hauteur
Grille.Move B_Left, B_Top
picBoutons.Top = Grille.Top + 200
fraOrigine.Top = Grille.Top
'Le dessin occupe la place restante
picTrajectoire.Move 100, 100, ScaleWidth - 200, B_Top - 200

txtDistance.Top = cmdFermer.Top - cmdFermer.Height
txtDistance.Left = txtOri(0).Left
lblDistance.Top = txtDistance.Top
lblDistance.Left = lblX.Left

' Traitement de recalcul des origines : ne doit ête fait que pour la feuille courante
If Not fCourante Is Me Then Exit Sub

' redéfinition du pt bas droit de l'image
xMaxPicture = picTrajectoire.ScaleWidth
yMaxPicture = picTrajectoire.ScaleHeight


'redéfinition du centre de l'image
neworx = picTrajectoire.ScaleWidth / 2
newory = picTrajectoire.ScaleHeight / 2
If gtOrx = neworx Then neworx = 0


' Nouvelles limites de visibilité ( en coordonnées réelles )
If neworx <> 0 And gtFacteurZoom <> 0 Then
  If FichierJournal Then Write #numFichLog, "Nouvelles limites de visibilité, numfille=", NumFille
  pmin.X = trEchelX(0, True)
  pmin.Y = trEchelY(yMaxPicture, True)
  pmax.X = trEchelX(xMaxPicture, True)
  pmax.Y = trEchelY(0, True)
  diagonalEcran = Sqr(Carre(pmax.X - gtMil.X) + Carre(pmax.Y - gtMil.Y))
' La nouvelle origine et le nouveau milieu ne seront actualisés qu'au prochain Rafraichir
' afin de pouvoir continuer à dessiner la trajectoire sans rafraichir
End If

'ScaleMode = savescale


End Sub


'**********************************************************************
'           Procédures relatives au dessin de la trajectoire

'           Dessiner
'           desSegment
'           desEchelle
'**********************************************************************


Public Sub dessiner(ByVal param As String, Optional ByVal chgtOrigine As Variant)

' param = "TOUT" ou "A_CALCULER"

Dim i As Integer, j As Integer
Dim Coul As Long
Dim memoNumEnreg
Dim objet As Object
Dim t As ligneTraj
Dim deplacement As Integer

  MDIGiration.MousePointer = vbHourglass  ' sablier
  memoNumEnreg = numEnreg
  Grille.OperationMode = 2      ' Mode Ligne en surbrillance
  ActivMenu Activation:=False
  
  On Error GoTo GestErr
  
  If param = "TOUT" Then           'dessin de toute la trajectoire
    If flagRecalcul Or flagExport Or pourEchelle Or Not IsMissing(chgtOrigine) Then
    Else
      RafraichirSimple memoNumEnreg
      Exit Sub
    End If
    
    Dim wLigne As Integer
    For wLigne = 0 To Grille.MaxRows - 1
      If Not TabLigneTraj(wLigne).Complet Then Exit For
      If wLigne = 0 Then ' Déplacement = DEBUT
        If flagExport Then
          InitialiseExport
        Else
          CalcDesCarross TabLigneTraj(wLigne), 0  ' superflu sauf en création ou chgt d'origine
        End If
              
      Else    '  wLigne>0 (déplacement <> DEBUT)
        activeCol wLigne + 1
           ' flagRecalcul : Retour options d'affichage ou Rafraichir après decompVirage
        If flagRecalcul Then TabLigneTraj(wLigne).dessinCalculé = False
        
        If TabLigneTraj(wLigne).dessinCalculé Then
          If IsMissing(chgtOrigine) Then
            If flagExport Then
            ' MODIF AV du 03/02/00 pour amélioration interne (les points à exporter sont déjà mémorisés)
              ExportCarrossRoues TabLigneTraj(wLigne)
            Else
              desSegment TabLigneTraj(wLigne), TabLigneTraj(wLigne - 1)
            End If
          
          Else
            effSegment TabLigneTraj(wLigne)
            desSegment TabLigneTraj(wLigne), TabLigneTraj(wLigne - 1)
          End If
          
        Else    ' Le segment n'a encore pas été calculé (lecture d'un .gir par exemple)
          desSegment TabLigneTraj(wLigne), TabLigneTraj(wLigne - 1)
        End If
      End If
    Next  ' wligne
    
  Else      ' param =  "A_CALCULER"
                'dessin du déplacement en cours, suivi du redessin des éventuels segments suivants
  
    For wLigne = numEnreg To Grille.MaxRows - 1 ' effacement des déplacements à partir du courant
      If Not TabLigneTraj(wLigne).Complet Then Exit For
      If Not TabLigneTraj(wLigne).dessinEffacé Then
        activeCol wLigne + 1
        effSegment TabLigneTraj(wLigne)
      End If
    Next wLigne
  
  '  desSegment NumEnreg     ' déplacement en cours
    activeCol numEnreg + 1
    desSegment TabLigneTraj(numEnreg), TabLigneTraj(numEnreg - 1)  ' déplacement en cours
    
    For numEnreg = memoNumEnreg + 1 To Grille.MaxRows - 1 ' redessin des déplacements suivants
      If Not TabLigneTraj(numEnreg).Complet Then Exit For
        activeCol numEnreg + 1
        
      If TabLigneTraj(numEnreg).dessinCalculé Then
        desSegment TabLigneTraj(numEnreg), TabLigneTraj(numEnreg - 1)
        ' réautorisation des champs qui auraient pu être interdits après insertion
        Grille.Col = 4  ' la longueur est en général tjs autorisée
        If Grille.Lock Then
          Grille.Lock = False
          deplacement = TabLigneTraj(numEnreg).deplacement
          Grille.Col = 3
          If deplacement <> LIGNE Then Grille.Lock = False
          Grille.Col = 5
          If deplacement = RACCORD Or deplacement = VIRAGE Then Grille.Lock = False
          If deplacement = RACCORD And Not TabLigneTraj(numEnreg + 1).Complet Then Grille.Col = 2: Grille.Lock = False
        End If
        
      Else    ' dessin non calculé
        desSegment TabLigneTraj(numEnreg), TabLigneTraj(numEnreg - 1)
      End If
    Next numEnreg    ' fin du re-dessin du segment
  
  End If  ' param (= "TOUT" ou "A_CALCULER")
    
  
  calculPositions
  
  If param = "TOUT" Then
    flagRecalcul = False
    chgEnreg memoNumEnreg + 1
  Else
    numEnreg = memoNumEnreg
  End If
  activeCol numEnreg + 1
  Grille.OperationMode = 0
  MDIGiration.MousePointer = vbDefault
  ActivMenu Activation:=True

Exit Sub

GestErr:
  
  Grille.OperationMode = 0
  MDIGiration.MousePointer = vbDefault
  ActivMenu Activation:=True
  ErrLocal Procédure:="Trajectoire.Dessiner"
  Err.Raise Err

End Sub

Private Sub RafraichirSimple(ByVal memoNumEnreg As Integer)
Dim i As Integer
Dim PremierPas

        ' Rafraichir simple : pas de recalcul à faire
       
   linMarque(0).Visible = False
   linMarque(1).Visible = False
   If fraOrigine.Visible Then PremierPas = 1 ' on ne redessine pas la première carrosserie en mise à jour de l'origine
   For i = PremierPas To maxPosition
     desPosition i
     If i Mod 10 = 1 Then DoEvents
   Next
  
   'NumEnreg = memoNumEnreg
   chgEnreg memoNumEnreg + 1
   activeCol numEnreg + 1
   Grille.OperationMode = 0
   MDIGiration.MousePointer = vbDefault
   ActivMenu Activation:=True

End Sub

Private Sub InitialiseExport()
Dim t As ligneTraj
Dim i As Integer, j As Integer
Dim pLine As POLYLINE

  If TabLigneTraj(1).deplacement = VIRAGE Then
    t = TabVirage(TabLigneTraj(1).numVirage).element(0)
  Else
    t = TabLigneTraj(1)
  End If
  ExportCarrossElem t, 0  ' Insertion des éléments de carrosserie du véhicule à l'origine
  
  ReDim DessinerRoues(0 To nbRemorque + 1) As Boolean
  'Indice 0 : Roues Avant du tracteur
  DessinerRoues(0) = OptionAffichage(NumFille).RouAvt.ADessiner
  'Indices 1 à nbRemorque+1 : Roues arrière de chaque élément du véhicule
  For i = 1 To nbRemorque + 1
    DessinerRoues(i) = OptionAffichage(NumFille).RouArr(i - 1).ADessiner
  Next
  
  With GraphExport.gphPolylines      ' Création des polylignes trace des roues
  'AV 16/09/03 : Correction sur les index pour PlanExport, MemoPtG et MemoPtD
    For j = 0 To nbRemorque + 1
      If DessinerRoues(j) Then
      ' Trace de la roue gauche
        Set pLine = New POLYLINE
        pLine.nomPlan = PlanExport(j + 3) 'PlanExport(3) = Roues Avant PlanExport(4 à 6) = Roues arrière de chaque élément de carosserie))
        .AddPline pLine
        .Item(.Count).Pts.Add t.MemoPtG(j, 0).X, t.MemoPtG(j, 0).Y
      ' Trace de la roue droite
        Set pLine = New POLYLINE
        pLine.nomPlan = PlanExport(j + 3)
        .AddPline pLine
        .Item(.Count).Pts.Add t.MemoPtD(j, 0).X, t.MemoPtD(j, 0).Y
      End If
    Next
  End With

  Set pLine = Nothing

End Sub

Private Sub calculPositions()
Dim i%, numpos%

' (re)Création du tableau des positions en vue du pas à pas ou d'un redessin simple (Zoom - Impressions ...)

Erase tabTrajPosition
Erase tabPtPosCarross
ReDim Preserve tabTrajPosition(0)
ReDim Preserve tabPtPosCarross(0)

For i = 1 To UBound(TabLigneTraj)
  With TabLigneTraj(i)
    If .dessinCalculé Then
      crePosition TabLigneTraj(i), i, numpos
    Else
      Exit For
    End If
  End With
Next i   ' tablignetraj
maxPosition = numpos

End Sub

Private Sub crePosition(t As ligneTraj, numtraj As Integer, numpos As Integer)
Dim i As Integer, j As Integer, intervalle As Integer
  ' Mémorisation dans le tableau tabTrajPosition du numéro de déplacement (tableau tabLigneTraj)
  ' Cas du VIRAGE : on attribue à chaque élément(numelem=0,1 ou2) du déplacement : -(numVirage*10 + numelem) Ex : -10,-11 et -12 si numVirage=1
  
  ' Mémorisation dans le tableau tabPtPosCarross du numéro de pt du tableau MemoPt correspondant au dessin des carrosseries
  ' Ces numéros sont espacés de 5 en 5 s'il y a des roues (+ un dernier éventuel si le nombre de pts du déplacement n'est pas un multiple d'intervalle(5))
  ' Ils sont espacés de 1 en 1 autrement
  ' Le point d'indice 0 du tableau MemoPt n'est pas mémorisé, car il est confondu avec le dernier point du déplacement précédent
  
    With t

      If .deplacement = VIRAGE Then
        With TabVirage(.numVirage)
          For i = 0 To 2
            crePosition .element(i), -(10 * t.numVirage + i), numpos
          Next
        End With
        Exit Sub
      End If

'      LgTot = LgTot + .Longueur
'      AngleTot = AngleTot + .DeltaBeta

      If Roues Then intervalle = 5 Else intervalle = 1

      For j = intervalle To UBound(.MemoPt, 2) Step intervalle
        numpos = numpos + 1
        ReDim Preserve tabTrajPosition(numpos)
        ReDim Preserve tabPtPosCarross(numpos)
        tabTrajPosition(numpos) = numtraj
        tabPtPosCarross(numpos) = j
      Next
      
      ' Ajout de l'éventuel dernier pt de carrosserie
      If j <> UBound(.MemoPt, 2) + intervalle Then
        numpos = numpos + 1
        ReDim Preserve tabTrajPosition(numpos)
        ReDim Preserve tabPtPosCarross(numpos)
        tabTrajPosition(numpos) = numtraj
        tabPtPosCarross(numpos) = UBound(.MemoPt, 2)
      End If
      
  End With  '  with t

End Sub

Public Sub desPosition(ByVal numpos As Integer)
' Dessin d'une position de carrosserie et des 5 positions de roues correspondantes
Dim i%, k%, numVirage%, numelem%, posCarross%
Dim t As ligneTraj
Static LigneActive%
  Dim intervalle%

  If Roues Then intervalle = 5 Else intervalle = 1
  
  k = tabTrajPosition(numpos)
  posCarross = tabPtPosCarross(numpos)
  
  If k <= 0 And numpos <> 0 Then ' Virage                         si k = -21 -->numvirage=2 et numelem=1
    numVirage = Abs(k) \ 10
    numelem = -k Mod 10
    t = TabVirage(numVirage).element(numelem)
    If numelem = 0 And posCarross = intervalle Then LigneActive = LigneActive + 1: activeCol LigneActive + 1
  Else
    'k = numéro de déplacement
    If LigneActive <> k Then LigneActive = k: activeCol k + 1
    t = TabLigneTraj(k)
  End If
  
  ' dessin des Carrosseries
  CalcDesCarross t, posCarross
  If Roues And numpos <> 0 Then
  ' Dessin des roues
      For i = (intervalle * ((posCarross - 1) \ intervalle)) + 1 To posCarross
         CalcDesRoues t, i, False
      Next
  End If

End Sub

Private Sub activeCol(ByVal Row As Integer)
          Grille.Row = Row
          Grille.Action = 0
'          DoEvents
End Sub

Public Sub effTout()
' Procédure appelée par le pas à pas
Dim i%
Dim FDPPresent As Boolean

    FDPPresent = fCourante.pminFDP.X <> 0 And fCourante.pmaxFDP.X <> 0
    If OptionAffichage(NumFille).FDP.ADessiner And FDPPresent Then
    
    ' Conserver le Fond de plan et effacer les éléments de trajectoire
      gtEffacement = True
      For i = 0 To maxPosition
        desPosition i
        DoEvents
      Next
      gtEffacement = False
    
    Else
    ' Effacer tout
      ObjetDessin.Cls
      Graduer
    End If
    
    ' Redessiner la 1ère position du tracteur
    CalcDesCarross TabLigneTraj(0), 0
  
End Sub

Private Sub effSegmentS(premier As Integer)
' Effacement de tous les segments à partir du numéro représenté par premier
Dim i%

  MDIGiration.MousePointer = vbHourglass
  Grille.OperationMode = 2      ' Mode Ligne en surbrillance
  For i = premier To Grille.MaxRows - 1
    activeCol i + 1
    effSegment TabLigneTraj(i)
  Next
  activeCol premier + 1
  Grille.OperationMode = 0
  MDIGiration.MousePointer = vbDefault

End Sub

Private Sub effSegment(t As ligneTraj)
Dim i%, effelem As Boolean

  If t.dessinEffacé Then Exit Sub
  If t.deplacement = VIRAGE Then
    With TabVirage(t.numVirage)
      effSegment .element(0)
      effSegment .element(1)
      effSegment .element(2)
    End With
    t.dessinEffacé = True
    Exit Sub
  End If
  
  gtEffacement = True
  
  With t
    For i = 1 To UBound(.MemoPt, 2)
      If Roues Then CalcDesRoues t, i, False
    Next
    
    For i = 1 To UBound(.MemoBeta, 2)
      effelem = Not Roues Or (i Mod 5 = 0) Or i = UBound(.MemoBeta, 2)
      If effelem Then
        CalcDesCarross t, i
        effelem = False
      End If
    Next
  
    .dessinEffacé = True
  
  End With

  gtEffacement = False

End Sub

Private Sub desSegment(t As ligneTraj, Tprec As ligneTraj)
' dessin du segment en cours
'------------------------------------------------------------------------------------------------------------------------------
' Cette procédure applique les formules de calcul telles qu'elles sont décrites dans le paragraphe
' Dessin de la trajectoire du Dossier de Programmation (p 8 et 10)
' Elle mémorise dans  tabMemoPt et tabMemoBeta les points et les angles effectivement utilisés par le dessin
' en vue d'une translation/rotation ultérieure sans recalcul complet des déplacements élémentaires
' Procédures appelées : calcPt pour calculer un pt par rapport au précédent
'                                       desEchelle pour dessin d'un segment avec mise à l'échelle
'------------------------------------------------------------------------------------------------------------------------------
Dim dalpha As Double, dbeta As Double, _
        alphad As Double, betad As Double, aSurV As Double
Dim pdeb As New PT, p As New PT, centre As New PT, pArr As New PT
Dim Coul As Long
Dim i%, nbpt As Integer
Dim pas As Single, lgReste As Single, deltaT As Single
Dim Calculer2Remorques As Boolean, Calculer1Remorque As Boolean

CalcDuree True
If t.deplacement = VIRAGE Then
  
  With TabVirage(t.numVirage)
    ' dessin des 3 déplacements élémentaires composant le VIRAGE
    desSegment .element(0), Tprec
    ' Mémorisation dans le déplacement VIRAGE de l'angle du 1er RACCORD
    For i = 0 To nbRemorque
      ReDim t.MemoBeta(i, 0)
    Next
    a = 0   '  pour CalcAngleRemorqueOptimise (élément CERCLE)
    desSegment .element(1), .element(0)
    a = -t.Vitbraq   '  pour CalcAngleRemorqueOptimise (élément RACCORD)
    desSegment .element(2), .element(1)
    ' Mémorisation dans le déplacement VIRAGE des valeurs du dernier RACCORD
    With .element(2)
      t.p.X = .p.X
      t.p.Y = .p.Y
      t.Beta = .Beta
      t.Sigma = .Sigma
      t.Gamma = .Gamma
    End With
  End With
  
  t.dessinCalculé = True
  t.dessinEffacé = False
  a = t.Vitbraq         ' car le signe de a a été inversé lors du dessin du dernier raccord
  Exit Sub
End If

If Roues Then             ' debug: si nbRemorque =2 and (carross(2) or roues(2))  -->  pas =ds
  pas = Sgn(v) * pas_roue
Else
  pas = Sgn(v) * pas_roue * 5
End If

    If nbRemorque > 0 Then
      With OptionAffichage(NumFille)
        Calculer1Remorque = .RouArr(1).ADessiner Or .Carross(1).ADessiner
        If nbRemorque = 2 Then
          Calculer2Remorques = .RouArr(2).ADessiner Or .Carross(2).ADessiner
        End If
      End With
    End If
    
  Coul = 0
  
  ' récupération du X,Y de la fin du déplacement précédent
  With Tprec
    pdeb.X = .p.X
    pdeb.Y = .p.Y
    If .deplacement = VIRAGE Then
      Alpha = 0     ' pour un virage, alpha correspond à l'élément cercle
    Else
      Alpha = .Alpha
    End If
    Beta = .Beta
    Sigma = .Sigma
    Gamma = .Gamma
    alphad = Alpha      ' utile à la remorque1
    betad = Beta         ' (fonctions CalcAngleRemorque et CalcAngleRemorqueOptimise)
  End With
  
  
With t
  a = .Vitbraq
  Erase .MemoPt
  Erase .MemoBeta
  Erase .MemoPtG
  Erase .MemoPtD
  nbpt = Fix(.Longueur / pas)       ' pas et longueur sont de même signe (signe de la vitesse)
  If nbpt * Abs(pas) < Abs(.Longueur) Then
    nbpt = nbpt + 1
  End If
  
  ReDim .MemoPt(nbRemorque, nbpt)
  ReDim .MemoBeta(nbRemorque, nbpt)
  ReDim .MemoPtG(nbRemorque + 1, nbpt)
  ReDim .MemoPtD(nbRemorque + 1, nbpt)
  .MemoPt(0, 0).X = pdeb.X    'tracteur
  .MemoPt(0, 0).Y = pdeb.Y
  .MemoBeta(0, 0) = Beta
  If nbRemorque > 0 Then
    .MemoPt(1, 0).X = pdeb.X - (Dpat(0) + e(0)) * Cos(Beta)
    .MemoPt(1, 0).Y = pdeb.Y - (Dpat(0) + e(0)) * Sin(Beta)
    .MemoBeta(1, 0) = Sigma
    If nbRemorque = 2 Then
      .MemoPt(2, 0).X = .MemoPt(1, 0).X - (Dpat(1) + e(1)) * Cos(Sigma)
      .MemoPt(2, 0).Y = .MemoPt(1, 0).Y - (Dpat(1) + e(1)) * Sin(Sigma)
      .MemoBeta(2, 0) = Gamma
    End If
    aSurV = a / v        ' pour CalcAngleRemorque et CalcAngleRemorqueOptimise
    
  End If

  If Roues Then CalcDesRoues t, 0, True       ' True <=> Recalcul complet  à faire
  
  Select Case .deplacement
  Case LIGNE
    dalpha = 0
    dbeta = 0
  Case CERCLE
    dalpha = 0
    dbeta = pas / e(0) * Sin(Alpha)
    pArr.X = pdeb.X - e(0) * Cos(Beta)
    pArr.Y = pdeb.Y - e(0) * Sin(Beta)
    Set centre = calcCentre(.Rayon, pArr)
    p.X = pdeb.X      ' le point suivant est calculé à partir du précédent
    p.Y = pdeb.Y
  Case RACCORD
    a = .Vitbraq
    deltaT = pas / v                ' deltaT est le temps mis par le véhicule pour faire un 'pas' - tjs positif (pas et v de même signe)
    dalpha = deltaT * a
    coefBeta = v / (e(0) * a)

  End Select
  
  p.X = pdeb.X
  p.Y = pdeb.Y

  For i = 1 To Fix(CSng(.Longueur / pas))  ' déplacement d'un incrément
      'Calcul du pt de référence (milieu de l'essieu avant) et de beta (et alpha si RACCORD)
    Select Case .deplacement
      Case RACCORD
        dbeta = -coefBeta * (Cos(Alpha + dalpha) - Cos(Alpha))
        calcPtRaccord p, deltaT, dalpha, dbeta
      Case CERCLE
          calcPtCercle p, centre, dbeta
      Case LIGNE
        calcPt p, pas
      End Select
    
    .MemoPt(0, i).X = p.X
    .MemoPt(0, i).Y = p.Y
    .MemoBeta(0, i) = Beta
    
    If Calculer2Remorques Then
      ' calcul de la 1è remorque avec la version non optimisée
        Set .MemoPt(1, i) = CalcAngleRemorque(p, alphad, betad, pas, aSurV) ' Ceci calcule en particulier sigma et gamma
        .MemoPt(2, i).X = .MemoPt(1, i).X + transRemorque(1) * Cos(Sigma)
        .MemoPt(2, i).Y = .MemoPt(1, i).Y + transRemorque(1) * Sin(Sigma)
        .MemoBeta(1, i) = Sigma
        .MemoBeta(2, i) = Gamma
    ElseIf Calculer1Remorque Then
      ' calcul de la 1è remorque avec la version optimisée
        Set .MemoPt(1, i) = CalcAngleRemorqueOptimise(p, alphad, betad, pas, aSurV) ' Ceci calcule en particulier sigma
        .MemoBeta(1, i) = Sigma
    End If
    
    
    If Not Roues Or (i Mod 5 = 0) Then
      CalcDesCarross t, i
    End If
    
    If Roues Then CalcDesRoues t, i, True
    
    pdeb.X = p.X
    pdeb.Y = p.Y
  
  Next i
  
    If i = nbpt Then ' calcul du segment résiduel  ( L n'est pas un multiple du pas)
      lgReste = .Longueur - (i - 1) * pas
      Select Case .deplacement
        Case RACCORD
        deltaT = lgReste / v                       ' deltaT est le temps mis par le véhicule pour faire la longueur restante - tjs positif (pas et v de même signe)
        dalpha = deltaT * a
        dbeta = -coefBeta * (Cos(Alpha + dalpha) - Cos(Alpha))
        calcPtRaccord p, deltaT, dalpha, dbeta
      Case CERCLE
        dbeta = lgReste / e(0) * Sin(Alpha)
        calcPtCercle p, centre, dbeta
      Case LIGNE
        calcPt p, lgReste
      End Select
    
      
      .MemoPt(0, i).X = p.X
      .MemoPt(0, i).Y = p.Y
      .MemoBeta(0, i) = Beta
      If Calculer2Remorques Then
      ' calcul de la 1è remorque avec la version non optimisée
          Set .MemoPt(1, i) = CalcAngleRemorque(p, alphad, betad, lgReste, aSurV) ' Ceci calcule en particulier sigma et gamma
          .MemoPt(2, i).X = .MemoPt(1, i).X - (Dpat(1) + e(1)) * Cos(Sigma)
          .MemoPt(2, i).Y = .MemoPt(1, i).Y - (Dpat(1) + e(1)) * Sin(Sigma)
          .MemoBeta(1, i) = Sigma
          .MemoBeta(2, i) = Gamma
      ElseIf Calculer1Remorque Then
      ' calcul de la 1è remorque avec la version optimisée
          Set .MemoPt(1, i) = CalcAngleRemorqueOptimise(p, alphad, betad, lgReste, aSurV) ' Ceci calcule en particulier sigma
          .MemoBeta(1, i) = Sigma
      End If
      
      CalcDesCarross t, i
      
      If Roues Then CalcDesRoues t, i, True

        
    Else    ' pas de traitement résiduel, on dessine une derniere carrosserie
            'sauf si çà vient juste d'être fait (roues=Vrai et i mod 5 = 1)
      
      If Roues And i Mod 5 <> 1 Then
        CalcDesCarross t, i - 1       'i-1 car i a été incrémenté à la fin de la boucle (qui tombait sur un nombre entier de pas)
      End If
      
    End If  ' traitement résiduel
    
    
  ' position en fin du déplacement
  .p.X = p.X
  .p.Y = p.Y
  
  
  .dessinCalculé = True
  .dessinEffacé = False
  .Beta = Beta          ' en principe inutile si le calcul est précis
  .Sigma = Sigma
  .Gamma = Gamma

End With
'  CalcDuree False

End Sub


Private Sub calcPt(ByVal p As PT, ByVal d As Single)
' calcul du pt p à une distance curviligne d de p1
' Référence: Dossier de programmation p.8
    
    p.X = p.X + d * Cos(Alpha + Beta)
    p.Y = p.Y + d * Sin(Alpha + Beta)

End Sub

Private Sub calcPtRaccord(ByVal p As PT, ByVal deltaT As Single, ByVal dalpha As Double, ByVal dbeta As Double)
 ' Référence Amélioration de la rapidité de la présentation (3/02/97)
 ' page 2
  p.X = p.X + coordRaccord(deltaT, True)
  p.Y = p.Y + coordRaccord(deltaT, False)
  Beta = Beta + dbeta
  Alpha = Alpha + dalpha

End Sub

Private Function coordRaccord(ByVal deltaT As Single, ByVal X As Boolean) As Double
' Calcul des coordoonnées du raccord (fonction appelée par calcPtRaccord)
' deltaT : temps de parcours correspondant au pas de représentation

Dim Somme, delta, IntN, w As Double
Dim epsilon As Double
Dim i%

epsilon = 0.001

delta = deltaT
Somme = (Fct(0, X) + Fct(deltaT, X)) / 2

Do
  IntN = Somme * delta
  delta = delta / 2
  w = delta
  While w < deltaT
    Somme = Somme + Fct(w, X)
    w = w + 2 * delta
  Wend
  i = i + 1
Loop While i <= 2 Or Abs(((Somme * delta) - IntN) / IntN) > epsilon

coordRaccord = IntN

End Function

Private Function Fct(ByVal t As Double, ByVal X As Boolean)
' Fonction caméléon(utilise le sinus ou le cosinus selon le booleen X) permettant de simplifier les formules
' dans la fonction CoordRaccord
Dim a1, a2 As Double

  a1 = Alpha + a * t
  a2 = a1 + Beta - coefBeta * (Cos(a1) - Cos(Alpha))
  If X Then Fct = v * Cos(a2) Else Fct = v * Sin(a2)
  
End Function

Private Function calcCentre(ByVal Rayon As Double, ByVal pEssArr As PT) As PT
' calcul du centre du cercle connaissant le rayon et le pt milieu de l'essieu arrière
Dim p As New PT, Trans As New PT, oscg As New PT

    Trans.X = -pEssArr.X
    Trans.Y = -pEssArr.Y
    Set oscg = Rotation(Trans, -Beta)
    p.X = 0 - oscg.X
    p.Y = Sgn(Alpha) * (Abs(Rayon) + voar(0)) - oscg.Y
    Set calcCentre = Rotation(p, Beta)
    
End Function

Private Sub calcPtCercle(ByVal p As PT, ByVal centre As PT, ByVal dbeta As Double)
Dim deltaX, deltaY As Double
' Calcul du nouveau pt milieu de l'essieu avant
' en appliquant au milieu de l'essieu arrière
' une rotation dbeta(constante sur tout le déplacement)

      deltaX = p.X - centre.X
      deltaY = p.Y - centre.Y
      ' rotation de l'essieu avant
      p.X = centre.X + deltaX * Cos(dbeta) - deltaY * Sin(dbeta)
      p.Y = centre.Y + deltaX * Sin(dbeta) + deltaY * Cos(dbeta)
      Beta = Beta + dbeta

End Sub

''Private Function calcPtCercle(ByVal p1 As PT, ByVal centre As PT, ByVal dbeta As Double) As PT
''Dim deltax, deltay As Double
''Dim p As New PT
' Calcul du nouveau pt milieu de l'essieu avant
' en appliquant au milieu de l'essieu arrière
' une rotation dbeta(constante sur tout le déplacement)

 ''     deltax = p1.X - centre.X
 ''     deltay = p1.Y - centre.Y
      ' rotation de l'essieu avant
''      p.X = centre.X + deltax * Cos(dbeta) - deltay * Sin(dbeta)
''      p.Y = centre.Y + deltax * Sin(dbeta) + deltay * Cos(dbeta)
''      Beta = Beta + dbeta
''      Set calcPtCercle = p

''End Function

Private Function CalcAngleRemorque(ByVal p As PT, alphad As Double, betad As Double, ByVal pas As Single, aSurV As Double) As PT
' Cette fonction calcule, par pas de 1cm (ds)  :
'        L'angle sigma de la remorque et la position p1 du point d'attache
'       L'angle gamma de la 2è Remorque
'       Les nouveaux alphad et betad
' p représente  la position du milieu de l'essieu avant du tracteur

Dim dalpha As Double, dbeta As Double, dsigma As Double
Dim Cpt As Integer
Dim p1 As New PT

' Référence : Dossier de programmation  - § Dessin de la trajectoire
  dalpha = ds * aSurV
  For Cpt = 1 To pas / ds
    dbeta = ds / e(0) * Sin(alphad) ' + dalpha)
    dsigma = (ds * Sin(alphad + betad - Sigma) + transRemorque(0) * dbeta * Cos(Sigma - betad)) / e(1)
    Gamma = Gamma + _
            (ds * Sin(alphad + betad - Gamma) + transRemorque(0) * dbeta * Cos(Gamma - betad) _
            + transRemorque(1) * dsigma * Cos(Sigma - Gamma)) / e(2)
    Sigma = Sigma + dsigma
    alphad = alphad + dalpha
    betad = betad + dbeta
  Next
  
  betad = Beta    ' pour corriger les erreurs d'arrondi dans la boucle
  p1.X = p.X - (Dpat(0) + e(0)) * Cos(Beta)
  p1.Y = p.Y - (Dpat(0) + e(0)) * Sin(Beta)
  
  Set CalcAngleRemorque = p1
        

End Function

Private Function CalcAngleRemorqueOptimise(ByVal p As PT, alphad As Double, betad As Double, ByVal SR As Single, ByVal aSurV As Double) As PT
' Cette fonction calcule, par intégration tous les 10 cm (10*ds)  :
'        L'angle sigma de la remorque et la position p1 du point d'attache (version optimisée : 1 seule remorque à représenter)
'       Les nouveaux alphad et betad
'  p représente  la position du milieu de l'essieu avant du tracteur

' Réf. : Dossier de programmation EPURE2.DOC du 05/02/97 corrigé par téléphone le 22/10/97 (erreur de signe sur FF dans l'algorithme)

Dim deltaS As Single, phi As Double
Dim sin_alpha As Double, cos_alpha As Double, sin_phi As Double, cos_phi As Double
Dim p1 As New PT
Dim fctA As Double, fctB As Double, Aprime As Double, Bprime As Double, dPhi As Double, dPhi2 As Double

' les coefficients K0Remorque et K1Remorque ont été calculés dans Form_Load
' K0Remorque =1/E(0) -  K1Remorque=Dpat(0)/E(0)/E(1)
' Phi : différentiel entre l'angle de la remorque et le tracteur  (beta - sigma)
' dPhi : dérivée première (notée fCourante dans le dossier)
'dPhi2 : dérivée seconde  (notée FF dans le dossier)
' fctA, fctB, Aprime et Bprime : notées A, B, A' et B' (pemettant de simplifier les notations de fCourante et FF)

sin_alpha = Sin(alphad)
cos_alpha = Cos(alphad)
phi = betad - Sigma

  Do
    deltaS = Min(Abs(ds * 10), Abs(SR)) * Sgn(v) 'calcul tous les 10 cm (ds*10) au lieu de 1cm
    SR = SR - deltaS
    
    sin_phi = Sin(phi)
    cos_phi = Cos(phi)

    fctA = K0Remorque + K1Remorque * cos_phi
    fctB = sin_phi / e(1)
    dPhi = fctA * sin_alpha - fctB * cos_alpha
    Aprime = -K1Remorque * sin_phi
    Bprime = cos_phi / e(1)
    dPhi2 = (aSurV * fctB + dPhi * Aprime) * sin_alpha + (aSurV * fctA - dPhi * Bprime) * cos_alpha
    
    phi = phi + dPhi * deltaS + dPhi2 * Carre(deltaS) / 2
    
    If a <> 0 Then
      alphad = alphad + aSurV * deltaS
      cos_alpha = Cos(alphad)
      sin_alpha = Sin(alphad)
    End If
  Loop While Abs(SR) > 0.000001
  
  betad = Beta
  Sigma = Beta - phi

  p1.X = p.X + transRemorque(0) * Cos(Beta)
  p1.Y = p.Y + transRemorque(0) * Sin(Beta)
  
  Set CalcAngleRemorqueOptimise = p1
 
End Function

Private Sub CalcDesCarross(t As ligneTraj, ByVal numpt As Integer)
' Calcul et Dessin de la (des) carrosserie(s)
' t: : ligne de déplacement
'numpt : numéro du pt du déplacement
Dim numRemorque As Integer

  With OptionAffichage(NumFille)
    For numRemorque = 0 To nbRemorque
      If .Carross(numRemorque).ADessiner Then FinDesCarross t, numpt, numRemorque
    Next
  End With
  
End Sub


Private Sub FinDesCarross(t As ligneTraj, ByVal numpt As Integer, ByVal numRemorque As Integer)
' Calcul et Dessin d'1 carrosserie (numRemorque : tracteur - remorque1 - remorque2)
' t: : ligne de déplacement
'numpt : numéro du pt du déplacement

  With Carross.gphInserts(numRemorque + 1)
    .p.X = t.MemoPt(numRemorque, numpt).X
    .p.Y = t.MemoPt(numRemorque, numpt).Y
    .rot = angConv(t.MemoBeta(numRemorque, numpt), unite)
    If Not (pourEchelle Or flagImpr) Then
      If Distance(.p.X, gtMil.X, .p.Y, gtMil.Y) >= maxDiagonal(numRemorque) + diagonalEcran Then Exit Sub
    End If
  End With
  
  DesCarross numRemorque
  
End Sub
Private Sub CalcDesRoues(t As ligneTraj, ByVal numpt As Integer, ByVal reCalcul As Boolean)
' Calcul et Dessin desRoues
' t: : ligne de déplacement
'numpt : numéro du pt du déplacement
Dim numRemorque%

  With OptionAffichage(NumFille)
    If .RouAvt.ADessiner Then CalcDesRouesAvant t, numpt, reCalcul
    For numRemorque = 0 To nbRemorque
      If .RouArr(numRemorque).ADessiner Then CalcDesRouesArr t, numpt, numRemorque, reCalcul
    Next
  End With

End Sub

Private Sub CalcDesRouesAvant(t As ligneTraj, ByVal numpt As Integer, ByVal reCalcul As Boolean)
' t: : ligne de déplacement
Dim p As New PT, Angle As Double
Dim objetPt As New PT

    With t
      If reCalcul Then
        Angle = .MemoBeta(0, numpt)
        Set .MemoPtG(0, numpt) = CalcRoue(.MemoPt(0, numpt), GAUCHE, Angle, numRemorque:=-1)     ' -1 : roues avant
        Set .MemoPtD(0, numpt) = CalcRoue(.MemoPt(0, numpt), DROITE, Angle, numRemorque:=-1)
      End If
    End With

    If numpt <> 0 Then DesRoues t, numpt, -1      ' -1 : roues avant
End Sub

  Private Sub CalcDesRouesArr(t As ligneTraj, ByVal numpt As Integer, ByVal numRemorque As Integer, reCalcul As Boolean)
' t: : ligne de déplacement
  Dim pArr As New PT, p As New PT, Angle As Double, objet As Object
  
    With t
      If reCalcul Then
        Angle = .MemoBeta(numRemorque, numpt)
        pArr.X = .MemoPt(numRemorque, numpt).X - e(numRemorque) * Cos(Angle)
        pArr.Y = .MemoPt(numRemorque, numpt).Y - e(numRemorque) * Sin(Angle)
        Set .MemoPtG(numRemorque + 1, numpt) = CalcRoue(pArr, GAUCHE, Angle, numRemorque)
        Set .MemoPtD(numRemorque + 1, numpt) = CalcRoue(pArr, DROITE, Angle, numRemorque)
      End If
    End With
      
    If numpt <> 0 Then DesRoues t, numpt, numRemorque

End Sub
        
'' Partie de DesRoues supprimée
''        If flagExport Then    CalcDesRoues n'est plus appelé si flagExport (03/02/2000)
''        ' ajout d'un point dans les polylignes dessinant les roues avant gauche et droite
''          With .MemoPtG(numRemorque+1, numpt)
''            p.X = .X
''            p.Y = .Y
''          End With
''          Set objetPt = GraphExport.Polylines(1).Pts.Add(p.X, p.Y)
''          With .MemoPtD(numRemorque+1, numpt)
''            p.X = .X
''            p.Y = .Y
''          End With
''          Set objetPt = GraphExport.Polylines(2).Pts.Add(p.X, p.Y)
''
''        ElseIf pourEchelle Then

Private Function CalcRoue(ByVal p As PT, ByVal côté As Integer, ByVal Angle As Double, ByVal numRemorque As Integer) As PT
' Calcul des positions gauche ou droite des roues
' au droit du pt p, milieu de l'essieu arrière de l' élément numRemorque de l'attelage (essieu avant du tracteur si numRemorque n'est pas renseigné)
' Angle repésente l'angle de l'émént de l'attelage concerné (beta,sigma ou gamma)
' numRemorque vaut de 0 à 2 pour les roues arrière - RouesAvant : numRemorque = -1
Dim p0 As New PT
Dim demilarg As Double

If numRemorque = -1 Then
  demilarg = voav
Else
  demilarg = voar(numRemorque)
End If

If côté = GAUCHE Then
  p0.X = p.X - demilarg * Sin(Angle)
  p0.Y = p.Y + demilarg * Cos(Angle)
Else
  p0.X = p.X + demilarg * Sin(Angle)
  p0.Y = p.Y - demilarg * Cos(Angle)
End If

Set CalcRoue = p0

End Function
  
 
 Private Sub DesRoues(t As ligneTraj, ByVal numpt As Integer, ByVal numRemorque As Integer)
 ' Dessin d'un segment de trace de roues
    'numRemorque=-1 pour les Roues avant, 0,1 ou 2 pour les roues arrière respectives
    
      If pourEchelle Then
      ' simple calcul mini/maxi
      With t.MemoPtG(numRemorque + 1, numpt)
        minimax .X, .Y
      End With
      With t.MemoPtD(numRemorque + 1, numpt)
        minimax .X, .Y
      End With
        
      Else
      ' dessin
        Dim couleur As Long
        If numRemorque = -1 Then  ' Roues avant
          couleur = OptionAffichage(NumFille).RouAvt.couleur
        Else
          couleur = OptionAffichage(NumFille).RouArr(numRemorque).couleur
        End If
        desEchelle t.MemoPtG(numRemorque + 1, numpt - 1), t.MemoPtG(numRemorque + 1, numpt), couleur
        desEchelle t.MemoPtD(numRemorque + 1, numpt - 1), t.MemoPtD(numRemorque + 1, numpt), couleur
      End If

End Sub

Private Sub ExportCarrossRoues(t As ligneTraj)
Dim Ok As Boolean

Dim numpt As Integer
Dim p As New PT
Dim objPol As POLYLINE
Dim i As Integer
ReDim wPass(2 + nbRemorque) As Boolean

If t.deplacement = VIRAGE Then
  With TabVirage(t.numVirage)
    ' Application aux 3 déplacements élémentaires composant le VIRAGE
    ExportCarrossRoues .element(0)
    ExportCarrossRoues .element(1)
    ExportCarrossRoues .element(2)
    Exit Sub
  End With

Else
  ' Polylignes Roues
  For Each objPol In GraphExport.gphPolylines
    If InstrRev(objPol.nomPlan, radicExport(3)) <> 0 Then
      i = 0 'Roues avant
    ElseIf InstrRev(objPol.nomPlan, radicExport(4)) <> 0 Then
      i = 1 'Roues arrière Tracteur
    ElseIf InstrRev(objPol.nomPlan, radicExport(5)) <> 0 Then
      i = 2 'Roues arrière Remorque
    ElseIf InstrRev(objPol.nomPlan, radicExport(6)) <> 0 Then
      i = 3 'Roues arrière Remorque2
    Else
      Err.Raise 654
    End If
      
    For numpt = 1 To UBound(t.MemoPt, 2)
    ' ajout des points du déplacement dans les polylignes dessinant les roues avant gauche ou droite
    ' la gauche ayant été créée en 1er dans 'Dessin', elle arrive en premier dans le For Each (indicateur wPass)
    ' on ne crée pas de point pour numpt=0, car il est confondu avec le dernier du déplacement précédent (Eventt l'élément DEBUT)
      If wPass(i) Then
        With t.MemoPtD(i, numpt)
          p.X = .X
          p.Y = .Y
        End With
      Else
        With t.MemoPtG(i, numpt)
          p.X = .X
          p.Y = .Y
        End With
      End If
      objPol.Pts.Add p.X, p.Y
    Next
    wPass(i) = True
  Next
  
  ' Carrosseries
  For numpt = 1 To UBound(t.MemoPt, 2)
    If Not Roues Or (numpt Mod 5 = 0) Then
      Ok = True
    Else
      Ok = (numpt = UBound(t.MemoPt, 2))
    End If
    If Ok Then ExportCarrossElem t, numpt
  Next
  
End If
  
End Sub

Private Sub ExportCarrossElem(t As ligneTraj, ByVal numpt As Integer)
Dim p As New PT
Dim rot As Double
Dim numRemorque As Integer
  
  For numRemorque = 0 To nbRemorque
    If OptionAffichage(NumFille).Carross(numRemorque).ADessiner Then
      p.X = t.MemoPt(numRemorque, numpt).X
      p.Y = t.MemoPt(numRemorque, numpt).Y
      rot = angConv(t.MemoBeta(numRemorque, numpt), unite)
      GraphExport.gphInserts.Add BlocExport(numRemorque), PlanExport(numRemorque), 256, "BYLAYER", p, 1, 1, rot, 1, 1, 0, 0
    End If
  Next

End Sub

Private Sub effRoues(t As ligneTraj, ByVal numpt As Integer)
' effacement des roues
Dim i%

    If OptionAffichage(NumFille).RouAvt.ADessiner Then
      desEchelle t.MemoPtG(0, numpt - 1), t.MemoPtG(0, numpt), 0
      desEchelle t.MemoPtD(0, numpt - 1), t.MemoPtD(0, numpt), 0
    End If
    
    For i = 0 To nbRemorque
      If OptionAffichage(NumFille).RouArr(i).ADessiner Then
        desEchelle t.MemoPtG(i + 1, numpt - 1), t.MemoPtG(i + 1, numpt), 0
        desEchelle t.MemoPtD(i + 1, numpt - 1), t.MemoPtD(i + 1, numpt), 0
      End If
    Next
    
End Sub

'**********************************************************************************
' Création (Recréation si pris en compte de surlargeur) des éléments de carrosserie
'**********************************************************************************
Public Sub CreCarross(Optional ByVal surLargeur As Boolean)
' Création des carrosseries de chaque élément de l'attelage
Dim rot As Double
Dim i%, numRemorque As Integer
Dim p As New PT

On Error GoTo GestErr

p.X = CDbl(txtOri(0))
p.Y = CDbl(txtOri(1))
rot = CDbl(txtOri(2))

For numRemorque = 0 To nbRemorque
  If Not surLargeur Then  ' Position correspondant à l'élément DEBUT
    With TabLigneTraj(0).MemoPt(numRemorque, 0)
      .X = p.X
      .Y = p.Y
    End With
  End If
  CreElemCarros numRemorque, p, rot, surLargeur
  
  If numRemorque < nbRemorque Then
    ' Décalage de la remorque par rapport à l'élément de carroserie précédent
    If nouveau Then
      p.X = p.X + transRemorque(numRemorque)
    Else  ' le décalage est déjà connnu (procédure 'lire')
      p.X = TabLigneTraj(0).MemoPt(numRemorque + 1, 0).X
      p.Y = TabLigneTraj(0).MemoPt(numRemorque + 1, 0).Y
      rot = TabLigneTraj(0).MemoBeta(numRemorque, 0)
    End If
  End If
Next

If Not surLargeur Then
  For i = 1 To 2
    Load linCarrosserie(i + 4 * (nbRemorque + 1))
  Next
  Set wTracteur = Carross.gphInserts(1)
End If

Exit Sub
GestErr:
  Err.Raise Err.Number
End Sub

'**********************************************************************************
' Création (Recréation si pris en compte de surlargeur) d'un élément de carrosserie
'**********************************************************************************
' Un élément de carroserie est inséré (point p)
'   Tracteur : Milieu de l'essieu avant
'   Remorque : Point d'atache de l'élément précédent

' avt : représente la distance entre le point d'insertion et l'avant de l'élément
' arr : représente la distance entre le point d'insertion et l'arrière de l'élément
' demilarg :  représente la demi-largeur de l'élément

Private Sub CreElemCarros(ByVal numRemorque As Integer, ByVal p As PT, ByVal rot As Double, ByVal surLargeur As Boolean)
Dim i As Integer
'Dim p1 As PT
Dim wBloc As New BLOC
'Dim objetPolyline As New POLYLINE, objetInsert As New INSERT
Dim arr As Double, avt As Double, demilarg As Double
Dim pLine As POLYLINE

On Error GoTo GestErr

  If Not surLargeur Then
    Set wBloc = Carross.gphBlocs.Add(Choose(numRemorque + 1, "Tracteur", "Remorque1", "Remorque2"))
    If numRemorque = TRACTEUR Then
      Set gtPlans = colPlans        ' gtPlans est déclaré dans GRAPHST.BAS
      For i = 0 To 7
        crePlanReserve i
      Next
    End If
  Else
    ' Suppresion des anciens contours de carrosserie
    Set wBloc = Carross.gphBlocs(numRemorque + 1)
    With wBloc.blkPolylines
        Do While .Count <> 0
          .Remove .Count
        Loop
    End With
  End If
  
  With wBloc
    arr = e(numRemorque) + PAFar(numRemorque)
    avt = PAFav(numRemorque)
    If numRemorque = TRACTEUR Then
      avt = avt + Debord(0)   ' AVANT
    End If
    
    If nbRemorque = numRemorque Then arr = arr + Debord(1)     ' Débordement arrière sur le dernier élément de carrosserie
    demilarg = larg(numRemorque) + SurLarg(numRemorque)       ' La surlargeur saisie correspond a priori à la demi-surlargeur (identique à droite et à gauche)
    maxDiagonal(numRemorque) = Max(Sqr(Carre(demilarg) + Carre(avt)), Sqr(Carre(demilarg) + Carre(arr)))
        
    Set pLine = New POLYLINE
    With pLine
      .nomPlan = GirPlanReserve(numRemorque)
      .ferme = True
      With .Pts
        .Add avt, demilarg
        .Add avt, -demilarg
        .Add -arr, -demilarg
        .Add -arr, demilarg
      End With
    End With
    .blkPolylines.AddPline pLine
    
    If numRemorque = TRACTEUR Then
      ' Création de la polyligne pour marquer le sens de la marche : visible si c'est demandé en option d'affichage (couleur du plan GirPlanReserve(7))
      Set pLine = New POLYLINE
      With pLine
        .nomPlan = GirPlanReserve(7)
        With .Pts
          .Add avt, 1
          .Add avt + 1, 0
          .Add avt, -1
        End With
      End With
      .blkPolylines.AddPline pLine
    End If
    
    If Not surLargeur Then
      ' Création de l'insert du tracteur correspondant à l'élément DEBUT
      Carross.gphInserts.Add .name, GirPlanReserve(numRemorque), 256, "BYLAYER", p, 1, 1, rot, 1, 1, 0, 0
      For i = 1 To 4: Load linCarrosserie(i + numRemorque * 4): Next
    End If
  End With    ' wBloc=Tracteur, Remoque1....

Exit Sub
GestErr:
  Err.Raise Err.Number
  
End Sub

Public Sub crePlanReserve(ByVal i As Integer)
Dim planNew As New PLAN
Dim ADessiner As Boolean
Static nomCourt As String

    If nbRemorque < 2 Then
      If i = 2 Or i = 6 Then Exit Sub
      If nbRemorque = 0 Then
        If i = 1 Or i = 5 Then Exit Sub
      End If
    End If
    
    With planNew
      If flagExport Then
        If nomCourt = "" Then nomCourt = Vehic.nomCourt
        PlanExport(i) = "GIR_" & nomCourt & radicExport(i)
        .name = PlanExport(i)
      Else
        .name = GirPlanReserve(i)
      End If
        
      Select Case i
      Case 0 To 2 ' Carrosseries
        .couleur = OptionAffichage(NumFille).Carross(i).couleur
        ADessiner = OptionAffichage(NumFille).Carross(i).ADessiner
      Case 3      ' Roues avant
        .couleur = OptionAffichage(NumFille).RouAvt.couleur
        ADessiner = OptionAffichage(NumFille).RouAvt.ADessiner
      Case 4 To 6 ' Roues arrière
        .couleur = OptionAffichage(NumFille).RouArr(i - 4).couleur
        ADessiner = OptionAffichage(NumFille).RouArr(i - 4).ADessiner
      Case 7      ' Nez du tracteur (si marquage du sens demandé)
        .couleur = OptionAffichage(NumFille).Carross(0).couleur
      End Select
      .typelign = "CONTINUOUS"
      
      On Error GoTo GestErr
      If flagExport Then
'        If ADessiner Then gtPlans.Add planNew, PlanExport(i)
        If ADessiner Then gtPlans.Add planNew
      Else
'        gtPlans.Add planNew, GirPlanReserve(i)
        gtPlans.Add planNew
        
      End If
    End With
    
Exit Sub

GestErr:  ' Plan déjà créé
  Resume Next

End Sub

Private Sub Grille_Advance(ByVal AdvanceNext As Boolean)
' Pour boucler dans la Grille , mais ne fonctionne pas

  With Grille
    If AdvanceNext Then
      .Row = 1
      .Col = 2
    Else
      .Row = .MaxRows
      .Col = 1
    End If
    .Action = 0
  End With
    
End Sub

Private Sub Grille_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
' bloc sélectionné dans la grille

  ligneSel = BlockRow
  ' Le click sur numéro de ligne a mis la valeur suivante à 7 anormalement au début de GrilleLeaveCell (bug du SPREAD)
  Grille.EditEnterAction = 5
  
End Sub

Private Sub Grille_Click(ByVal Col As Long, ByVal Row As Long)
  ' debug : bug du SPREAD: si on clique dans le menu
  ' et qu'une donnée a été saisie, on plantera lors de la fermeture de frmTraject
  ' le pendant de cette instruction est dans barre_Click dans MDIGiration
  
''''  Grille.EditModePermanent = True

Exit Sub
If Col = 1 Then
  Grille.Col = Col
  Grille.Row = Row
  Grille.BackColor = bleupale
End If

End Sub

Private Sub Grille_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If Grille.IsBlockSelected Then
    Select Case KeyCode
    Case vbKeyInsert
        InserDepl
    Case vbKeyDelete
        SupprDepl
    End Select
  End If
  
End Sub

Private Sub Grille_LostFocus()

' On ignore le traitement qui va suivre
' Mais celui-ci peut faire modifier la colonne active, car en sortie grille.col est tjs égal à 5, et non à la valeur en cours
 Grille.Col = Grille.activeCol
  
End Sub

'**************************************************************************
' Clic bouton droit sur la grille : Menu contextuel Insérer-Supprimer-Décomposer
'**************************************************************************
Private Sub Grille_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

  If Button = vbRightButton Then
    If Grille.IsBlockSelected And ligneSel <> -1 Then
    Else
      Grille_BlockSelected -1, Grille.ActiveRow, -1, Grille.ActiveRow
    End If
      Grille.Row = ligneSel
      Grille.Row2 = ligneSel
      Grille.Col = -1
      Grille.Col2 = -1
      Grille.Action = 2
    
    On Error GoTo GestErr
    With MDIGiration
      For i = 4 To .mnuOutils.Count
        .mnuOutils(i).Visible = False
      Next
      If TabLigneTraj(ligneSel - 1).deplacement = VIRAGE Then
        PopupMenu .mnuBarre(1), , , , .mnuOutils(MNUDECOMPVIRAGE)
      Else
        PopupMenu .mnuBarre(1), , , , .mnuOutils(MNUINSERDEPL)
      End If
      For i = 4 To .mnuOutils.Count
        .mnuOutils(i).Visible = True
      Next
    End With
  End If
  
  Exit Sub
  
GestErr:
  If Err <> 9 Then
    ErrGeneral
  Else
    Exit Sub
  End If
  
End Sub

'**********************************************************************
'           Procédures relatives à la saisie de la Grille

'       Grille_Change
'       Grille_LeaveCell
'       verifCompatibDepl
'**********************************************************************

Private Sub Grille_Change(ByVal Col As Long, ByVal Row As Long)
  
Dim i%
Dim flagInsert As Boolean
' verifCompatibDepl peut retourner flagInsert=True, indiquant que l'on n'est as en mode Ajout

If Col <> 1 Then
  TabLigneTraj(numEnreg).Modif = True
  Exit Sub
End If

  If Not TabLigneTraj(Row - 2).Complet Then
    MessA IDm_PrecedentIncomplet
    With Grille
      .Row = Row
      .Col = Col
      .TypeComboBoxCurSel = -1
      numEnreg = Row - 2
      .Row = Row - 1
      chgEnreg (.Row)   ' chgenreg modife .Row
      .Row = Row - 1
      For i = 2 To 5
        .Col = i                'repositionnement sur la 1ère colonne autorisée
        If .Lock = False Then Exit For
      Next
      .Action = 0
      .SetFocus
    End With
    Exit Sub
  End If
  
    ' saisie du type de déplacement
    ' si compatible avec le précédent, alimentation des valeurs fixes et verrouillage des cellules non modifiables
    ' Référence: Dossier de programmation p4 et suivantes - 'Définition des déplacements élémentaires
  
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect
  
  With Grille
    .Col = Col          ' = 1
    .Row = Row
    If verifCompatibDepl(.TypeComboBoxCurSel, flagInsert) Then
      .Col = Col
      .TypeComboBoxCurSel = TabLigneTraj(numEnreg).deplacement
      .SetFocus
      
    Else
'      La ligne qui suit devrait être réactivée si on autorise de réinitialiser une ligne complète
'      If TabLigneTraj(NumEnreg).Complet Then effsegment TabLigneTraj(numenreg)
      .Col = Col
      TabLigneTraj(numEnreg).deplacement = .TypeComboBoxCurSel
    
      Select Case .TypeComboBoxCurSel
      Case LIGNE
        If FichierJournal Then Write #numFichLog, "Ligne " & Row & " : LIGNE"
        For i = 2 To 5
          .Col = i
          If i = 4 Then
            .Lock = False
            .BackColor = blanc
          End If
        Next
        With TabLigneTraj(numEnreg)
          .Rayon = 0
          .DeltaBeta = 0
          .Vitbraq = 0
          .Alpha = 0
          .Beta = TabLigneTraj(numEnreg - 1).Beta
        End With
      
      Case CERCLE
        If FichierJournal Then Write #numFichLog, "Ligne " & Row & " : CERCLE"
        .Col = 2
        .Value = AffGrille(TabLigneTraj(numEnreg - 1).Rayon)
        .Col = 3
        .Lock = False
        .BackColor = blanc
        .Col = 4
        .Lock = False
        .BackColor = blanc
        .Col = 5
        With TabLigneTraj(numEnreg)
          .Rayon = TabLigneTraj(numEnreg - 1).Rayon
          .Vitbraq = 0
          .Alpha = TabLigneTraj(numEnreg - 1).Alpha
          Alpha = .Alpha
        End With
        betadeb = TabLigneTraj(numEnreg - 1).Beta
      
      Case RACCORD
        If FichierJournal Then Write #numFichLog, "Ligne " & Row & " : RACCORD"
        With TabLigneTraj(numEnreg - 1)
          If .deplacement = VIRAGE Then ' dans un virage, R représente le rayon sur le cercle, et non le rayon de fin de déplacement
            Alphadeb = TabVirage(.numVirage).element(2).Alpha
          Else
            Alphadeb = .Alpha
          End If
          betadeb = .Beta
        End With
        
        For i = 2 To 5
          .Col = i
          If i = 2 And flagInsert Then Else .Lock = False
          .BackColor = blanc
        Next
        
        ' si la ligne suivante est déjà typée suite à insertion:
        If UBound(TabLigneTraj) > numEnreg Then
          If TabLigneTraj(numEnreg + 1).deplacement = CERCLE Then
            .Col = 2
            .Lock = True
            Alpha = TabLigneTraj(numEnreg + 1).Alpha
            R = R_alpha("alpha")
            TabLigneTraj(numEnreg).Alpha = Alpha
            TabLigneTraj(numEnreg).Rayon = R
            TabLigneTraj(numEnreg).DernierChampSaisi(1) = 1
          End If
        End If

      Case VIRAGE
        If FichierJournal Then Write #numFichLog, "Ligne " & Row & " : VIRAGE "
        With TabLigneTraj(numEnreg - 1)
          Alphadeb = 0    ' en principe 0
          betadeb = .Beta
        End With
        For i = 3 To 5
          .Col = i
          .Lock = False
          .BackColor = blanc
        Next
        TabLigneTraj(numEnreg).numVirage = nbVirage
        ReDim Preserve TabVirage(nbVirage)
        With TabVirage(nbVirage)
          .element(0).deplacement = RACCORD
          .element(1).deplacement = CERCLE
          .element(2).deplacement = RACCORD
          .element(1).Vitbraq = 0
          .element(2).Rayon = 0
          .element(2).Alpha = 0     ' en principe
        End With
        nbVirage = nbVirage + 1
      End Select
        
        
      If numEnreg = .MaxRows - 1 Then ' Création d'une nouvelle ligne
        creLigne
      Else
        With TabLigneTraj(.MaxRows - 1)
          .DernierChampSaisi(1) = 0
          .Complet = False
        End With
      End If
    
    End If
  End With  ' Grille
      
End Sub

Private Sub Grille_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

' --------------------
' Sortie d'une cellule
' --------------------

' _
        Avant de changer de cellule, vérification de la cohérence de la saisie _
        Si OK : mémorisation du dernier champ saisi _
                                        si la ligne est complète, (re)dessiner l'épure _
                                        interdiction du champ déplacement _
        sinon : réinitialisation du champ _
                                        refus du changement de cellule ( Cancel=True ) _
        ------------------------------------------------------------------------------
                                        
Dim i%, Dernier%, numsuiv%, numdepl%
Dim Valeur As Variant, ChampEnCours As Integer
Dim Deltabeta1 As Double, Deltabeta2 As Double, L1 As Double, L2 As Double
Dim deltaBetaInconnu As Boolean, aInconnu As Boolean, LInconnu As Boolean, reponse
Dim dessinFait As Boolean


On Error GoTo GestErr

  If NewRow = -1 Then Exit Sub                     ' Focus hors de la Grille
  
  
  With TabLigneTraj(numEnreg)
    If Col = 1 Or Not .Modif Then
      If Row <> NewRow Then chgEnreg (NewRow)
      If NewCol = 1 Then Grille.EditEnterAction = 7 Else Grille.EditEnterAction = 5 ' Touche RETURN=TAB ou non
      Exit Sub                                ' Le type de déplacement est géré par la procédure Grille_Change
    End If
  
    .Modif = False
    Grille.Col = Col
    Grille.Row = Row
    ChampEnCours = Col - 1
    
    If Grille.Value = String(Len(Grille.Value), " ") Then Grille.Value = ""
    If Grille.Value = "" Then
    ' Traitement d'une cellule remise à blanc : possible dans certains cas uniquement pour RACCORD ou VIRAGE
    '------------------------------------------------------------------------------------------------------------------------------------------------------
      If .deplacement = RACCORD Then
'        If MsgBox("Réinitialiser la ligne", vbOKCancel + vbExclamation + vbDefaultButton2) = vbCancel Then
         MessConfirm IDm_ReinitLigne, vbOKCancel '       , vbOKCancel + vbExclamation + vbDefaultButton2) = vbCancel Then
'        End If
      End If
      
      Select Case .deplacement
      Case VIRAGE, RACCORD
        If .deplacement = VIRAGE Then
          Select Case ChampEnCours
          Case RAY
            .Rayon = Empty
          Case ANGGIR
            .DeltaBeta = Empty
          Case LG
            .Longueur = Empty
          Case VBRAQ
            .Vitbraq = Empty
          End Select
          Grille.Col = 2  ' rayon
          Grille.Value = ""
          If .DernierChampSaisi(1) = ChampEnCours Then   ' le dernier champ saisi est celui qu'on efface
            .DernierChampSaisi(1) = .DernierChampSaisi(0)
          End If
        
        Else      ' RACCORD
          Grille.Col = 2  ' rayon
          If Grille.Lock Then   '  le déplacement suivant a déjà été saisi: rayon non modifiable
            .DernierChampSaisi(1) = 1
          Else
            .Rayon = Empty
            Grille.Value = ""
            .DernierChampSaisi(1) = 0
          End If
              ' Réinitialisation des 3 autres champs
           .DeltaBeta = Empty: .Longueur = Empty: .Vitbraq = Empty
          For i = 3 To 5: Grille.Col = i: Grille.Value = "": Next
          .DernierChampSaisi(0) = 0
        End If
        
      .Complet = False
      ' La réinitialisation de la ligne implique son effacement et des suivantes
        If Not .dessinEffacé Then effSegmentS numEnreg
        .dessinCalculé = False
        If ExportOK Then ExportOK = False: GriserMenus Etat:=True
        If TabLigneTraj(numEnreg + 1).deplacement = -1 Then Grille.Col = 1: Grille.Lock = False
      
        Grille.Col = Col
        Grille.Action = 0
        Cancel = True
        Grille.SetFocus
        Exit Sub
      
      Case Else   ' Remise à blanc interdit pour DEBUT,LIGNE et CERCLE
        Err.Raise 110
      End Select
    End If    ' Grille.value = ""
  End With  ' tablignetraj(numenreg)
  
   
  
  '---------------------------- vérif de 1er niveau -------
  With Grille
    
    If numEnreg <> 0 Then
      If Not TabLigneTraj(numEnreg - 1).Complet Then
        Mess1 IDm_PrecedentIncomplet
      End If
    End If
    
    If Not IsNumeric(.Value) Then
      Mess1 IDm_Numerique
    
    ElseIf Col = 3 And Abs(angConv(.Value, radian)) > 2 * pi Then
      Mess1 IDm_AngleMaxi & CStr(angConv(2 * pi, unite))
    
    ElseIf (Col = 4 Or Col = 5) And CDbl(.Value) = 0 Then
      Mess1 IDm_DifferentDeZero
    
    ElseIf Col = 3 And TabLigneTraj(numEnreg).deplacement <> RACCORD And CDbl(.Value) = 0 Then    ' zéro autorisé pour angle de giration uniqt pour un raccord (courbe en S)
      Mess1 IDm_DifferentDeZero
    
    ElseIf Col = 4 Then
      If CDbl(.Value) < 0 Then
        Mess1 ID_Longueur & " " & ID_Strict & " " & ID_PositifLongueur
      ElseIf CDbl(.Value) / pas_roue > nbIteration Then
        MessConfirm IDm_LgDepltImportant & vbCrLf & IDm_Confirmation, vbYesNo
      End If
    
    ElseIf Col = 5 And Superieur(Abs(angConv(CDbl(.Value), radian)), aVehMax) Then
      Mess2 IDm_VitBraqTropGrand, ID_Maximum, ID_Admissible, angConv(aVehMax, unite), -1
    
    ElseIf Col = 2 And TabLigneTraj(numEnreg + 1).deplacement <> -1 Then 'rayon
      numsuiv = TabLigneTraj(numEnreg + 1).deplacement
      If (CDbl(.Value = 0) Or CDbl(.Value) >= INFINI) And numsuiv = CERCLE Then
        Mess1 ID_ValeurRayon & " : " & Minus(IDm_SuivantIncompatible)
      ElseIf CDbl(.Value) <> 0 And CDbl(.Value) < INFINI And (numsuiv = LIGNE Or numsuiv = VIRAGE) Then
        Mess1 ID_ValeurRayon & " : " & Minus(IDm_SuivantIncompatible)
      End If
    
    End If
  
   
  '---------------------------- Traitement -------
  
      Valeur = CDbl(.Value)
      If FichierJournal Then Write #numFichLog, "Ligne " & Row, libCol(Col - 2), Valeur
      
      With TabLigneTraj(numEnreg)
      Select Case .deplacement
    
      Case DEBUT
    '-------------------
        R = CDbl(Valeur)
        Alpha = R_alpha("rayon")
        .Alpha = Alpha
        .Rayon = R
        .Complet = True
        
        Grille.Row = Row + 1
        With TabLigneTraj(numEnreg + 1)
          If .deplacement = CERCLE Then
            R = R_alpha("alpha")   ' pour réafficher le nouveau rayon
            Grille.Row = Row
            .Rayon = R
            .Alpha = Alpha
        ElseIf Not IsEmpty(.Rayon) Then    'ce cas ne peut arriver que pour un Raccord dont seul le rayon a été saisi
          Grille.Col = 2
          Grille.Value = ""
          .Rayon = Empty
          .DernierChampSaisi(1) = 0
        End If
      End With
      Grille.Row = Row
 
      Case LIGNE
    '------------------
        l = CDbl(Valeur) * Sgn(v)
        .Longueur = l
        Grille.Col = 4
        Grille.Value = AffGrille(Abs(l))
        .Complet = True
       
      Case CERCLE
   '----------------------
        
        With Grille
        Select Case ChampEnCours
      
        Case 2  'angle de giration
          DeltaBeta = angConv(CDbl(Valeur), radian)
          l = L_deltabeta("anggir")
'          If Not Cancel And L / pas_roue > nbIteration Then
'            Cancel = (MsgBox("Longueur de déplacement importante" & Chr(13) & "Confirmation ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo)
'          End If
          
        Case 3  ' longueur
          l = CDbl(Valeur) * Sgn(v)
          DeltaBeta = L_deltabeta("longueur")
      
        End Select  ' case ChampEnCours
        
        End With
        
          .Longueur = l
          .DeltaBeta = DeltaBeta
          Beta = betadeb + DeltaBeta
          .Beta = Beta
          .Complet = True
      
      Case RACCORD
   '-------------------------

        If .DernierChampSaisi(1) <> ChampEnCours Then
          Dernier = .DernierChampSaisi(1)
        Else
          Dernier = .DernierChampSaisi(0)
        End If
      
      If FichierJournal Then Write #numFichLog, ChampEnCours, Dernier
      
        Select Case ChampEnCours
      
        Case RAY    'Rayon
        '_______________
        
          R = CDbl(Valeur)
          Alpha = R_alpha("rayon")

        Select Case Dernier
        
        Case ANGGIR    ' angle de giration
          a = a_deltabeta("anggir", ChampEnCours)
          l = L_a("vitbraq")
        
        Case LG    ' longueur
          a = L_a("longueur", ChampEnCours)
          DeltaBeta = a_deltabeta("vitbraq")
        
        Case VBRAQ    ' vitesse de braquage
          verifSigne_a_dAlpha ChampEnCours

          DeltaBeta = a_deltabeta("vitbraq")
          l = L_a("vitbraq")
        End Select       'case dernier
            
  ' Alimentation du tableau TabLigneTraj
        .Alpha = Alpha
        .Rayon = R
        .Complet = .DernierChampSaisi(1) <> 0 And Dernier <> 0
        If .Complet Then
          Beta = betadeb + DeltaBeta
          .Beta = Beta
        End If
      
      Case ANGGIR  ' angle de giration
      '__________________________
      
        If CDbl(Valeur) = 0 And Alphadeb = 0 Then
          Mess1 IDm_2RayonsInfinisConsecutifs
        End If
        
        DeltaBeta = angConv(CDbl(Valeur), radian)
        Beta = Beta + betadeb
        
        Select Case Dernier
        '                       ---------
        Case RAY  ' rayon
          Alpha = R_alpha("rayon")
          a = a_deltabeta("anggir", ChampEnCours)
          l = L_a("vitbraq")
        
        Case LG                  ' longueur
          Alpha = calcAlpha("", ChampEnCours)
          a = L_a("longueur", ChampEnCours)
          R = R_alpha("alpha")
        
        Case VBRAQ  ' vitesse de braquage
          Alpha = calcAlpha("anggir", ChampEnCours)
          l = L_a("vitbraq")
          R = R_alpha("alpha")
        
        End Select    'case dernier
        
        
  ' Alimentation du tableau TabLigneTraj
        .DeltaBeta = DeltaBeta
        .Beta = Beta
        .Complet = .DernierChampSaisi(1) <> 0 And Dernier <> 0
        If .Complet Then                                'debug: test rajouté dans la version SPREAD (6.02.97)
          .Alpha = Alpha
        End If
        
      Case LG ' longueur
      '_______________
      
        l = CDbl(Valeur) * Sgn(v)
        
        Select Case Dernier
        '                       ---------
        
        Case RAY           ' rayon
          Alpha = R_alpha("rayon")
          a = L_a("longueur", ChampEnCours, Dernier)
          DeltaBeta = a_deltabeta("vitbraq")
        
        Case ANGGIR          ' angle de giration
          Alpha = calcAlpha("", ChampEnCours)
          a = L_a("longueur", ChampEnCours, Dernier)
          R = R_alpha("alpha")
          
        Case VBRAQ          ' vitesse de braquage
          Alpha = calcAlpha("longueur", ChampEnCours)
          DeltaBeta = a_deltabeta("vitbraq")
          R = R_alpha("alpha")
          
        End Select          ' case dernier
        
        
  ' Alimentation du tableau TabLigneTraj
        .Longueur = l
        .Complet = .DernierChampSaisi(1) <> 0 And Dernier <> 0
        If .Complet Then
          .Alpha = Alpha
          Beta = betadeb + DeltaBeta
          .Beta = Beta
        End If
              
      Case VBRAQ ' vitesse de braquage
      '___________________________
      
        a = angConv(CDbl(Valeur), radian)
        
        Select Case Dernier
        '                       ---------

        Case RAY  ' rayon
          Alpha = R_alpha("rayon")
          verifSigne_a_dAlpha ChampEnCours
          l = L_a("vitbraq")
          DeltaBeta = a_deltabeta("vitbraq")
          
         Case ANGGIR         ' angle de giration
          Alpha = calcAlpha("anggir", ChampEnCours)
          l = L_a("vitbraq")
          R = R_alpha("alpha")
          
        Case LG          ' longueur
          Alpha = calcAlpha("longueur", ChampEnCours)
          DeltaBeta = a_deltabeta("vitbraq")
          R = R_alpha("alpha")
        End Select     'case dernier
        
        
  ' Alimentation du tableau TabLigneTraj
        .Vitbraq = a
        .Complet = .DernierChampSaisi(1) <> 0 And Dernier <> 0
        If .Complet Then
          .Alpha = Alpha
          Beta = betadeb + DeltaBeta
          .Beta = Beta
        End If
            
      End Select  ' case dernier
      
      ' Dernier controle de verrouillage : ne devrait jamais survenir - évite un plantage au cas où (Réf .ENV2506.DOC)
      If .Complet And (Superieur(Abs(Alpha), alphaVehMax) Or Superieur(Abs(a), aVehMax) Or l * v < 0 Or (Alpha - Alphadeb) * a < 0) Then
        Mess1 IDm_ErrConstr
      End If
        
      If .Complet Then  ' debug: a priori superflu
        .Rayon = R
        .DeltaBeta = DeltaBeta
        .Longueur = l
        .Vitbraq = a
      End If
      
      If TabLigneTraj(numEnreg + 1).deplacement = CERCLE Then
        Grille.Row = Row + 1
        R = R_alpha("alpha")   ' pour réafficher le nouveau rayon
        Grille.Row = Row
        With TabLigneTraj(numEnreg + 1)
          .Rayon = R
          .Alpha = Alpha
        End With
      End If
      
      Case VIRAGE
   '-------------------------
    
      Dernier = .DernierChampSaisi(1)
      Grille.Col = 5 ' vitesse de braquage
      aInconnu = (Grille.Value = "")
      Grille.Col = 4 ' longueur
      LInconnu = (Grille.Value = "")
      Grille.Col = 3 ' DeltaBeta
      deltaBetaInconnu = (Grille.Value = "")
      
      
      Select Case ChampEnCours
      
      Case ANGGIR  ' angle de giration
      '_____________________
        DeltaBeta = angConv(CDbl(Valeur), radian)
        If Not LInconnu And aInconnu Then ' longueur
          verifDeltabetaL
        End If
        If Not aInconnu And LInconnu Then ' vitesse de braquage
          verifSigne_a_Deltabeta
        End If
          
    
      Case LG ' longueur
      '______________
        l = CDbl(Valeur) * Sgn(v)
        If Not deltaBetaInconnu And aInconnu Then ' angle de giration
          verifDeltabetaL
        End If
    
      Case VBRAQ ' vitesse de braquage
      '_______________________
        a = angConv(CDbl(Valeur), radian)
        If Not deltaBetaInconnu And LInconnu Then ' angle de giration
          verifSigne_a_Deltabeta
        End If

      End Select  ' case col-1
      
      If Dernier <> 0 Then
        .Complet = True
        For i = 3 To 5
          If i <> Col Then
            Grille.Col = i
            If .Complet Then .Complet = Grille.Value <> ""
          End If
        Next
      
      End If
      
      If .Complet Then
                     ' Virage complet: vérifications de cohérence
                     
                     ' verification du signe de a et deltabeta
        verifSigne_a_Deltabeta ChampEnCours
        
        CalculElemVirage Deltabeta1, L1, Deltabeta2, L2, ChampEnCours
          
          R = R_alpha("alpha")
          .Rayon = R
          .Alpha = Alpha
          .Beta = betadeb + DeltaBeta
          With TabVirage(.numVirage)
            With .element(0)
              .Rayon = R
              .DeltaBeta = Deltabeta1
              .Longueur = L1
              .Vitbraq = a
              .Alpha = Alpha
              .Beta = betadeb + Deltabeta1
              .DernierChampSaisi(1) = 1    ' En cas de décomposition, le RACCORD, s'il est modifié, sera calculé à partir du rayon et de  la valeur saisie
            End With
            With .element(1)
              .Rayon = R
              .DeltaBeta = Deltabeta2
              .Longueur = L2
              .Vitbraq = 0
              .Alpha = Alpha
              .Beta = betadeb + Deltabeta1 + Deltabeta2
            End With
            With .element(2)
              .DeltaBeta = Deltabeta1
              .Longueur = L1
              .Vitbraq = -a
              .Beta = betadeb + DeltaBeta
              .DernierChampSaisi(1) = 1    ' En cas de décomposition, le RACCORD, s'il est modifié, sera calculé à partir du rayon et de  la valeur saisie
            End With
          End With
              

        .DeltaBeta = DeltaBeta
        .Longueur = l
        .Vitbraq = a
      End If   ' Virage complet
        
    End Select  'case type de deplacement
    
    End With
    

'-------------  Traitement final ----------------------------------

      With TabLigneTraj(numEnreg)
        If FichierJournal Then If .Complet Then Write #numFichLog, "OK - Complet" Else Write #numFichLog, "OK - Incomplet"
        If .DernierChampSaisi(1) <> ChampEnCours Then
            ' champ saisi différent du précédent
          If TabLigneTraj(numEnreg + 1).Complet And .deplacement = RACCORD Then
            ' raccord bloqué : le rayon est tjs considéré comme le dernier saisi
            .DernierChampSaisi(0) = ChampEnCours
          Else
            .DernierChampSaisi(0) = .DernierChampSaisi(1)
            .DernierChampSaisi(1) = ChampEnCours
          End If
        End If
        
        ' dessiner si possible
        If .Complet And numEnreg <> 0 And Not TabLigneTraj(Max(0, numEnreg - 1)).dessinEffacé Then
          dessiner "A_CALCULER"
          dessinFait = True
          Modif = True
          VerifExport
        Else
            ' Il est possible que le déplacement ait été calculé dans une action précédente, mais effacé pour raison diverse (insertion - remise à blanc d'un champ de saisie au-dessus)
          TabLigneTraj(numEnreg).dessinCalculé = False
        End If
      End With
                        
      If dessinFait Then ' Bloquage du type de déplacement (si on le réautorise, voir message dans grilleChange)
        .Row = Row
        .Col = 1
        .Lock = True
        numdepl = TabLigneTraj(Row - 2).deplacement
        If numdepl = RACCORD Or numdepl = DEBUT Then
          .Row = Row - 1    ' bloquage du Rayon du déplacement précédent
          .Col = 2
          .Lock = True
          .Row = Row
          With TabLigneTraj(Row - 2)
            If .DernierChampSaisi(1) <> 1 Then
              .DernierChampSaisi(0) = .DernierChampSaisi(1)
              .DernierChampSaisi(1) = 1
            End If
          End With
        End If
        Grille.Col = NewCol
        activeCol NewRow         ' car l'appel à la fonction  'dessin' a fait balader le curseur
      End If    ' dessinFait
      
      If Row <> NewRow Or dessinFait Then                       ' changement de ligne
        chgEnreg (NewRow)                   ' il faut tjs réinitialiser les variables globales si appel du module dessin
      End If
      
      If NewCol = 1 Then Grille.EditEnterAction = 7 Else Grille.EditEnterAction = 5
  
  End With    ' Grille

Exit Sub


GestErr:

  If Err = 110 Then
    Cancel = True
    retrouveLigne Col
    Exit Sub
  ElseIf Err = 111 Then
    Cancel = True
  Else
    ErrGeneral
  End If
    
End Sub
  
Function verifCompatibDepl(ByVal numdepl As Integer, flagInsert As Boolean) As Boolean

    ' Vérification de la compatibilité du déplacement avec le précédent puis éventuellement avec le suivant
    ' Référence: Dossier de programmation p7
    ' flagInsert indique en retour que l'on n'est pas en mode Ajout
    
    Dim rein As Boolean
      
On Error GoTo GestErr

  If TabLigneTraj(numEnreg).DernierChampSaisi(1) <> 0 Then
    MessConfirm IDm_ReinitLigne, vbOKCancel
  End If
  
  rein = (TabLigneTraj(numEnreg).deplacement <> -1)
  flagInsert = False
  
With TabLigneTraj(numEnreg - 1)
  If .deplacement = VIRAGE Then
    R = 0
  Else
    R = .Rayon
  End If
  Select Case numdepl
  Case LIGNE
    If R <> 0 Then
      Mess1 IDm_PrecedentIncompatible
    ElseIf .deplacement = LIGNE Then
      Mess1 IDm_2LignesConsecutives
    End If
  
  Case RACCORD
    
  Case CERCLE
    If R = 0 Then
      Mess1 IDm_PrecedentIncompatible
    ElseIf .deplacement = CERCLE Then
      Mess1 IDm_2CerclesConsecutifs
    ElseIf .deplacement = VIRAGE Then
      Mess1 IDm_PrecedentIncompatible
    End If
  
  Case VIRAGE
    If R <> 0 Then ' Cercle ou raccord de rayon fini
      Mess1 IDm_PrecedentIncompatible
    End If

  End Select      ' numdepl
  
  End With          ' tabLigneTraj(Numenreg-1)
  

    If Grille.MaxRows > numEnreg + 1 Then ' Suite à  insertion d'une ligne
    
      With TabLigneTraj(numEnreg + 1)
          Alphadeb = .Alpha - .Longueur * .Vitbraq / v    ' Inversion de la formule du Dossier de progr.  p.6 "RACCORD" (R et deltabeta)
          If Alphadeb <> 0 Then
            If (e(0) / Tan(Alphadeb) - Sgn(Alphadeb) * voar(0) >= INFINI) Then Alphadeb = 0
          End If
          
        Select Case numdepl
        Case LIGNE
          If .deplacement = LIGNE Then
            Mess1 IDm_2LignesConsecutives
          ElseIf .deplacement = CERCLE Then
            Mess1 IDm_SuivantIncompatible
          ElseIf .deplacement = RACCORD And Abs(Alphadeb) > 0.00000000001 Then
            Mess1 IDm_ContinuiteRayonImpossible
          End If
        
        Case CERCLE
          If .deplacement = LIGNE Then
            Mess1 IDm_SuivantIncompatible
          ElseIf .deplacement = CERCLE Then
            Mess1 IDm_2CerclesConsecutifs
          ElseIf .deplacement = VIRAGE Then
            Mess1 IDm_SuivantIncompatible
          ElseIf .deplacement = RACCORD Then
            If Alphadeb = 0 Then
              Mess1 IDm_ContinuiteRayonImpossible
            ElseIf Abs(Alphadeb - TabLigneTraj(numEnreg - 1).Alpha) > 0.00000000001 Then
              Mess1 IDm_ContinuiteRayonImpossible
            End If
          End If
      
        Case VIRAGE
          If .deplacement = CERCLE Then
            Mess1 IDm_SuivantIncompatible
          ElseIf .deplacement = RACCORD And Abs(Alphadeb) > 0.00000000001 Then
            Mess1 IDm_ContinuiteRayonImpossible
          End If
          
        Case RACCORD
          If .deplacement = LIGNE Or .deplacement = VIRAGE Then
            If TabLigneTraj(numEnreg - 1).deplacement = LIGNE Or TabLigneTraj(numEnreg - 1).deplacement = VIRAGE Then
                Mess1 IDm_ContinuiteRayonImpossible
            ElseIf Abs(Alphadeb - TabLigneTraj(numEnreg - 1).Alpha) < Abs(angConv(0.001 * .Longueur / v, radian)) Then
              Mess1 IDm_SuivantIncompatible & Chr(13) & IDm_RayonFinalProcheDepart
            Else
              Alpha = 0
              R = R_alpha("alpha")
              flagInsert = True
              TabLigneTraj(numEnreg).Rayon = R
              TabLigneTraj(numEnreg).Alpha = Alpha
              TabLigneTraj(numEnreg).DernierChampSaisi(1) = 1 ' Le raccord sera calculé à partir du rayon (forcément bloqué pour compatibilité)
            End If
          ElseIf .deplacement = CERCLE Then
            R = .Rayon
            Alpha = R_alpha("rayon")
            If Abs(Alpha - TabLigneTraj(numEnreg - 1).Alpha) < Abs(angConv(0.001 * .Longueur / v, radian)) Then
              Grille.Value = ""
              Mess1 IDm_SuivantIncompatible & Chr(13) & IDm_RayonFinalProcheDepart
            Else
              Grille.Value = AffGrille(.Rayon)
              flagInsert = True
              TabLigneTraj(numEnreg).Rayon = R
              TabLigneTraj(numEnreg).Alpha = Alpha
            End If
          ElseIf .deplacement = RACCORD Then
            If Abs(Alphadeb - TabLigneTraj(numEnreg - 1).Alpha) < Abs(angConv(0.001 * .Longueur / v, radian)) Then
              Mess1 IDm_SuivantIncompatible & Chr(13) & IDm_RayonFinalProcheDepart
            Else
              Alpha = Alphadeb
              R = R_alpha("alpha")
              flagInsert = True
              TabLigneTraj(numEnreg).Rayon = R
              TabLigneTraj(numEnreg).Alpha = Alpha
              TabLigneTraj(numEnreg).DernierChampSaisi(1) = 1      ' Le raccord sera calculé à partir du rayon (forcément bloqué pour compatibilité)
            End If
          End If
        
        End Select    ' numdepl
      End With        ' TabLigneTraj(NumEnreg + 1)
    
    End If

    
  If rein Then
    reinLigne numEnreg
    Grille.SetFocus ' Le focus a été perdu suite à boite de dialogue
  End If

  Exit Function
  
GestErr:
    
    If Err = 110 Then
      verifCompatibDepl = True
    Else
      ErrGeneral
    End If
    
End Function

Private Sub autorChamp()
' Autrisation de l'accès aux cellules selon le type de déplacement
Dim i%

  With Grille
      .Col = 1
      Select Case .TypeComboBoxCurSel
      Case LIGNE
          .Col = 4
          .Lock = False
      
      Case CERCLE
        .Col = 3
        .Lock = False
        .Col = 4
        .Lock = False
      
      Case RACCORD
        For i = 2 To 5
          .Col = i
          .Lock = False
        Next
       
      Case VIRAGE
        For i = 3 To 5
          .Col = i
          .Lock = False
        Next
      End Select
    End With
        

End Sub

Private Sub retrouveLigne(ByVal Col As Integer)
' retrouver les anciennes valeurs d'une ligne suite à saisie erronée

Dim Valeur As Variant


      With TabLigneTraj(numEnreg)
        Select Case Col
        Case 2
          If IsEmpty(.Rayon) Then
            Valeur = ""
          ElseIf .Rayon = 0 Then
            Valeur = ID_INFINI
            R = 0
          Else
            Valeur = .Rayon
            R = Valeur
          End If
          
        Case 3
          If IsEmpty(.DeltaBeta) Then
            Valeur = ""
          Else
            Valeur = angConv(.DeltaBeta, unite)
            DeltaBeta = .DeltaBeta
          End If
          
        Case 4
          If IsEmpty(.Longueur) Then
            Valeur = ""
          Else
            Valeur = Abs(.Longueur)
            l = Valeur
          End If
          
        Case 5
          If IsEmpty(.Vitbraq) Then
            Valeur = ""
          Else
            Valeur = angConv(.Vitbraq, unite)
            a = .Vitbraq
          End If
        End Select
      
        If .deplacement = VIRAGE Then
          If Not .dessinCalculé Then .Complet = False
        End If
      End With
      
      With Grille
        .Col = Col
        If VarType(Valeur) = vbString Then
        .Value = Valeur
      Else
        .Value = AffGrille(Valeur)
      End If
      .SetFocus   ' nécessaire, car on a probablement cliqué dans une boite de dialogue
    End With

End Sub

Private Sub chgEnreg(ByVal NewRow As Integer)
  
  numEnreg = NewRow - 1
  With TabLigneTraj(numEnreg)
    R = .Rayon
    DeltaBeta = .DeltaBeta
    l = .Longueur
    a = .Vitbraq
    Alpha = .Alpha
    Beta = .Beta
    
    If .deplacement = VIRAGE And Not .Complet Then
    ' pour les virages, les valeurs ne sont mémorisées dans la table qu'après validation totale
      Grille.Row = NewRow
      Grille.Col = 3
      If Grille.Value <> "" Then DeltaBeta = angConv(Grille.Value, radian)
      Grille.Col = 4
      If Grille.Value <> "" Then l = Grille.Value
      Grille.Col = 5
      If Grille.Value <> "" Then a = angConv(Grille.Value, radian)
    End If
       
  End With
  
  AfficheMarque TabLigneTraj(numEnreg), 0
  
  If numEnreg <> 0 Then
    With TabLigneTraj(numEnreg - 1)
      If .deplacement = VIRAGE Then ' dans un virage, R représente le rayon sur le cercle, et non le rayon de fin de déplacement
        Alphadeb = TabVirage(.numVirage).element(2).Alpha
      Else
        Alphadeb = .Alpha
      End If
      betadeb = .Beta
      AfficheMarque TabLigneTraj(numEnreg - 1), 1
    End With
  Else
    linMarque(1).Visible = False
  End If
  
  AfficheAngleTot
  
End Sub

Private Sub AfficheMarque(ByRef t As ligneTraj, ByVal i As Integer)
 Dim largeurMarque As Single
  
  ' On n'affiche pas les marques en saisie de l'origine
  ' Il ne faut pas non + appeler cette fonction depuis l'imprimante, le grand dessin ou l'export
 If fraOrigine.Visible Or Not ObjetDessin Is picTrajectoire Or flagExport Or pourEchelle Then Exit Sub
 
 linMarque(i).Visible = OptionAffichage(NumFille).position
 If gtFacteurZoom = 0 Or Not OptionAffichage(NumFille).position Then Exit Sub
 
  largeurMarque = Max(2 * larg(0), 300 / gtFacteurZoom)
  With t
    linMarque(i).X1 = trEchelX(.p.X + largeurMarque * Cos(.Beta + pi / 2), toReel:=False)
    linMarque(i).Y1 = trEchelY(.p.Y + largeurMarque * Sin(.Beta + pi / 2), toReel:=False)
    linMarque(i).X2 = trEchelX(.p.X - largeurMarque * Cos(.Beta + pi / 2), toReel:=False)
    linMarque(i).Y2 = trEchelY(.p.Y - largeurMarque * Sin(.Beta + pi / 2), toReel:=False)
    linMarque(i).Visible = (.dessinEffacé = False)
  End With
End Sub

Private Sub AfficheAngleTot(Optional ByVal nbEnreg As Integer = -1)
Dim i As Integer
Dim Break As Boolean
Dim AngleTot As Single

  lblAngleTotal = ""
  AngleTot = 0#
  
  If nbEnreg = -1 Then nbEnreg = numEnreg + 1
  
  For i = 1 To nbEnreg - 1
    With TabLigneTraj(i)
      If .Complet Then
        AngleTot = AngleTot + .DeltaBeta
      Else
        Break = True
        Exit For
      End If
    End With
  Next
  
  lblAngleTotal = ID_AngleTotal & " = " & AffGrille(angConv(AngleTot, unite)) & " " & libUnite(OptGen.unite)
  lblAngleTotal.Enabled = Not Break
      
End Sub

Private Sub creLigne()
' création d'une ligne supplémentaire dans la Grille
Dim i%
  
If VersionDemo And Grille.MaxRows = MAXLIGNEDEMO + 1 Then Exit Sub

  With Grille
    ReDim Preserve TabLigneTraj(.MaxRows)
    TabLigneTraj(.MaxRows).deplacement = -1
    TabLigneTraj(.MaxRows).dessinEffacé = True
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    For i = 2 To 5: .Col = i: .Lock = True: .BackColor = grisPale: Next
      If VersionDemo And .MaxRows > MAXLIGNEDEMO Then
        .Col = 1
        .Lock = True
        .CellType = 1   ' Texte
        .Value = ID_FinProjet
      End If
  End With

End Sub

Private Sub reinLigne(ByVal numligne As Integer)
' réinitialisation d'une ligne, suite à
'     changement de type de déplacement
'     insertion
'     suppression sans recollement des lignes

' numligne est tjs égal à NumEnreg : il est passé en paramètre pour une meilleure lisiblité

Dim i%

  With TabLigneTraj(numligne)
    .deplacement = -1
    .Rayon = Empty
    .DeltaBeta = Empty
    .Longueur = Empty
    .Vitbraq = Empty
    Set .p = Nothing
    .DernierChampSaisi(1) = 0
    .Complet = False
    If ExportOK Then ExportOK = False: GriserMenus Etat:=True
    .dessinCalculé = False
    .dessinEffacé = True
    Erase .MemoPt
    Erase .MemoBeta
  End With
  
  With Grille
    .Row = numligne + 1
    For i = 2 To 5: .Col = i: .Value = "": .Lock = True: .BackColor = grisPale: Next
    .Col = 1
    .Lock = False
'    .BackColor = bleupale
    .Action = 0   ' activation du champ déplacement
  End With
  
End Sub


'************************************************************
'
' Les fonctions ci-dessous sont utilisées par la fonction Grille_LeaveCell
'
' Les formules correspondent au Dossier de programmation p.5 à 7

'************************************************************

Private Sub CalculElemVirage(Deltabeta1 As Double, L1 As Double, Deltabeta2 As Double, L2 As Double, ByVal ChampEnCours As Integer)
' Décomposition des différents éléments du  Virage :
' L1 et DeltaBeta1 = longueur et angle de chaque raccord
' L2 et Deltabeta2 = longueur et angle du cercle
'ChampEnCours représente le champ en cours de saisie (et de validation) : Angle de giration, longueur ou vitesse de braquage
' Réf. Dossier de programmation p.5

Dim alphaMax As Double, deltaBetaTotal As Double
Dim k As Double, Lmin As Double, aMin As Double, kMaxi As Double
Dim alphaMini As Double, alphaMaxi As Double, wAlpha As Double, pAlpha As Double

        
        
        ' Calcul aux limites
        
        alphaMax = (a / 2) * (l / v)          ' correspond à un cercle de longueur nulle
        If Abs(alphaMax) > alphaVehMax Then alphaMax = Sgn(alphaMax) * alphaVehMax    ' limitation supplémentaire
        deltaBetaTotal = SommeDeltaBeta(alphaMax, Deltabeta1, L1, Deltabeta2, L2)
        
        If Abs(deltaBetaTotal) < Abs(DeltaBeta) Then
            Select Case ChampEnCours
            Case 2 ' Angle de giration
'              Mess2 IDm_Reduire & Id_Amplitude & ID_DeLAngle, ID_Maximum, "", Abs(angConv(deltaBetaTotal, unite)), -1
              Mess2 IDm_Reduire & ID_AmplitudeAngle, ID_Maximum, "", Abs(angConv(deltaBetaTotal, unite)), -1
            Case 3 ' Longueur
'        Réf. FAX du 16/05/97
              k = 2 * Abs(v / a)
              If e(0) * Abs(DeltaBeta) - k * (1 - Cos(alphaVehMax)) > 0 Then
                Lmin = LgVirage(alphaVehMax, k, Abs(DeltaBeta))
              Else
                Lmin = LgVirage(Arccos(1 - e(0) * Abs(DeltaBeta) / k), k, Abs(DeltaBeta))
              End If
              Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Abs(Lmin), 1
            Case 4 ' Vitesse de braquage
              kMaxi = (e(0) * Abs(DeltaBeta) - Abs(l) * Sin(alphaVehMax)) / ((1 - Cos(alphaVehMax)) - alphaVehMax * Sin(alphaVehMax))
              If kMaxi > Abs(l) / alphaVehMax Then
                    ' dichotomie  Réf: FAX du 14/05/97 p2
                If a < 0 Then   ' car alpha (braquage sur le cercle) est forcément du signe de a
                  alphaMini = -alphaVehMax
                  alphaMaxi = 0
                  wAlpha = alphaMini
                Else
                  alphaMini = 0
                  alphaMaxi = alphaVehMax
                  wAlpha = alphaMaxi
                End If
                pAlpha = -e(0) * DeltaBeta / l    'Pente du segment [alpha,alphaD]  (pour un virage, alphaD=0)

                While Abs(alphaMaxi - alphaMini) > 0.000001
                  If pAlpha < pente(wAlpha) Then alphaMini = wAlpha Else alphaMaxi = wAlpha   ' la fonction pente utilise alphadeb, qui est bien égal à 0 pour un virage
                  wAlpha = (alphaMini + alphaMaxi) / 2
                Wend
                kMaxi = l / wAlpha
              End If
              
              If Abs(kMaxi) < Abs(2 * v / aVehMax) Then
                Mess1 IDm_ConstructionImpossible
              Else
                aMin = 2 * v / kMaxi
                Mess2 IDm_AugmenterVitBraq, ID_Minimum, "", Abs(angConv(aMin, unite)), 1
              End If
            End Select
          
          
        Else      ' Somme des DeltaBeata OK
        
        ' Calcul effectif des différents composants
          alphaMaxi = Sgn(v) * Abs(alphaMax)
          alphaMini = -alphaMaxi
          Alpha = alphaMax
          While Abs(deltaBetaTotal - DeltaBeta) > 0.000001
            If DeltaBeta < deltaBetaTotal Then alphaMaxi = Alpha Else alphaMini = Alpha
            Alpha = (alphaMini + alphaMaxi) / 2
            deltaBetaTotal = SommeDeltaBeta(Alpha, Deltabeta1, L1, Deltabeta2, L2)
          Wend
        End If

        End Sub

Private Function LgVirage(ByVal Alpha As Double, ByVal k As Double, ByVal DeltaBeta As Double) As Double
'Calcul de la longueur d'un Virage pour deltabeta et a donnés
'        Réf. FAX du 16/05/97

    LgVirage = (e(0) * DeltaBeta - k * (1 - Cos(Alpha))) / Sin(Alpha) + k * Alpha

End Function

Private Function R_alpha(ByVal param As String) As Variant
' calcul du rayon en fonction de alpha ou inversement
    
    Grille.Col = 2
    If param = "rayon" Then             'calcul de alpha
      If R = 0 Or R >= INFINI Or R <= -INFINI Then
        R = 0
        R_alpha = 0
        Grille.Value = ID_INFINI             ' réécriture du rayon
      Else
        If Abs(R) < RVehMin Then
          Mess2 IDm_Augmenter & ID_AmplitudeRayon, ID_Minimum, ID_Absolu, RVehMin, 1
        End If
        R_alpha = Atn(e(0) / (R + Sgn(R) * voar(0)))    ' Dossier de progr. p4 § "ORIGINE-DEBUT" et p.6 "RACCORD"
        Grille.Value = AffGrille(R)
      End If
      
      If TabLigneTraj(numEnreg).deplacement = RACCORD And R_alpha = Alphadeb Then  ' Réf. Dossier de progr. p6 §conditions générales
        Mess1 IDm_PrecedentIncompatible
      End If
      
    Else                                             'calcul de R
      If Alpha = 0 Then
        R_alpha = 0     ' ou R_alpha=INFINI?????
        Grille.Value = ID_INFINI
      Else
        R_alpha = e(0) / Tan(Alpha) - Sgn(Alpha) * voar(0)      ' Dossier de progr. p3 et p,5 § "CERCLE-1"
        Grille.Value = AffGrille(R_alpha)
      End If
    End If
    
End Function

Private Function L_deltabeta(ByVal param As String)
' calcul de la longueur en fonction de l'angle de giration ou inversement
' pour un CERCLE cf Dossier de prog. p5 CERCLE-2

    If param = "anggir" Then
      L_deltabeta = e(0) * DeltaBeta / Sin(Alpha)
      If Sgn(L_deltabeta) = Sgn(v) Then
        Grille.Col = 4
        Grille.Value = AffGrille(Abs(L_deltabeta))
      Else
        Mess1 IDm_IncompatibleRayonSigneAngle  ' Demande MO d 23/04/97
      End If
      
    Else
      L_deltabeta = l / e(0) * Sin(Alpha)
      Grille.Col = 3
      Grille.Value = AffGrille(angConv(L_deltabeta, unite))
    End If

End Function

Private Function a_deltabeta(param As String, Optional ByVal ChampEnCours As Variant) As Variant
 ' calcul de la vitesse de braquage en fonction de l'angle de giration ou inversement
' Si param="anggir", ChampEnCours représente le champ en cours de saisie (et de validation) : rayon ou angle de giration
Dim Rayon(1) As Double

        If param = "anggir" Then
          
          ' 1er Controle : Deltabeta=0 si et seulement si  Rdeb=-Rfin (Courbe en S)
          If DeltaBeta = 0 Or R = -TabLigneTraj(numEnreg - 1).Rayon Then
                                        '   Rectificatif du 24/04/97  - p1 et 2 haut (cas R|deltabeta et deltabeta|R)
            If DeltaBeta = 0 And R = -TabLigneTraj(numEnreg - 1).Rayon Then
              Mess1 IDm_DepltIndetermine
            Else
              Mess1 IDm_ConstructionImpossible
            End If
' Fin du rectificatif
        End If
        
          ' 1er controle OK   ( Deltabeta est forcément <>0)
          
          If ChampEnCours = RAY Then      ' rayon
            ' Réf: messages d'erreur RACCORD (14/02/97)  p1bas (Deltabeta|R)
                  ctrlDeltabetaR
          Else  ' ChampEnCours=angle de giration
              If v / DeltaBeta * (Cos(Alphadeb) - Cos(Alpha)) / (Alpha - Alphadeb) <= 0 Then
              ' Réf: messages d'erreur RACCORD (14/02/97)  p2 haut (R|Deltabeta)
                Mess1 IDm_ChangerAngleDeSigne
              End If
          End If
          
              ' calcul de la vitesse de braquage
          a_deltabeta = v / (e(0) * DeltaBeta) * (Cos(Alphadeb) - Cos(Alpha))
              
              ' Controle des bornes : (aVehMin=0.0005) < a < aVehMax
              '-----------------------------------------------------------------------------
            If Abs(angConv(a_deltabeta, unite)) < aVehMin Then
              Mess1 IDm_VitBraqTropFaibleInduite
            ElseIf Abs(a_deltabeta) > aVehMax Then
                                    ' Réf. Rectificatif du 24/04/97    - p2 - cas R|deltabeta
              If ChampEnCours = ANGGIR Then    ' angle de giration
                Mess2 IDm_Augmenter & ID_AmplitudeAngle, ID_Minimum, "", Abs(angConv(v / (e(0) * aVehMax) * (Cos(Alphadeb) - Cos(Alpha)), unite)), 1
              Else    ' ChampEnCours = Rayon
                                            ' Complément au Dossier de programmation du 14/02/97
                If Alphadeb = 0 Then
                  Mess1 IDm_Augmenter & ID_AmplitudeRayon
                Else
                  Rayon(0) = e(0) / Tan(Alphadeb) - Sgn(Alphadeb) * voar(0)            ' Dossier de progr. p3 (§ Symboles utilisés)
                  Rayon(1) = e(0) / Tan(alphaVehMax) - alphaVehMax * voar(0)           ' Dossier de progr. p3 (§ Symboles utilisés)
                  If Sgn(Alphadeb) <> Sgn(v / DeltaBeta) Then       ' Correspond dans le Complément au Dossier à Lim1 et Lim2 de même signe
                    Mess1 IDm_Augmenter & ID_AmplitudeRayon & vbCrLf & IDm_SansDepasser & Format(Abs(Rayon(0)), "#.00")
                  ElseIf Sgn(R) = Sgn(Rayon(0)) Then
                    Mess1 IDm_Reduire & ID_EcartRayonPrecedent & "(" & Format(Rayon(0), "#.00") & ")" & vbCrLf & IDm_OuChangerDeSigne
                  Else
                    Mess1 IDm_Reduire & ID_AmplitudeRayon & vbCrLf & IDm_OuLeChangerDeSigne
                  End If
                End If
                
              End If
            End If
          
          Grille.Col = 5
        
        Else     ' param="vitbraq""
          a_deltabeta = v / (e(0) * a) * (Cos(Alphadeb) - Cos(Alpha))
          Grille.Col = 3
          
        End If
        
      ' afichage de la valeur calculé au forma de la grille
        Grille.Value = AffGrille(angConv(a_deltabeta, unite))
        
End Function

Private Sub ctrlDeltabetaR()
  ' Référence Complément au Dossier de programmation 14/02/97 (erreur.doc) - Messages d'erreur des raccords p1 et 2(haut)
  ' Fonction appelée par a_deltabeta (param = "anggir" ChampEnCours=rayon), avant le calcul de la vitesse de braquage
  ' pour vérifier que le rayon saisi est bien compatible avec l'angle déjà connu DeltaBeta)
   
Dim k, alphaLim1, alphaLim2 As Double, RLim1 As Double, RLim2 As Double
Dim Rayon(1) As Double
Dim Erreur As Boolean

k = v / DeltaBeta
alphaLim1 = Min(-Alphadeb, Sgn(k) * alphaVehMax)
alphaLim2 = Max(-Alphadeb, Sgn(k) * alphaVehMax)

' En valeur absolue, RLim1 correspond à RVehMin(=R(alphavehmax)) et RLim2 à R(-alphadeb) ou inversement

If alphaLim1 = 0 Then
  RLim1 = 0
Else
  RLim1 = e(0) / Tan(alphaLim1) - Sgn(alphaLim1) * voar(0)
End If
If alphaLim2 = 0 Then
  RLim2 = 0
Else
  RLim2 = e(0) / Tan(alphaLim2) - Sgn(alphaLim2) * voar(0)
End If

Erreur = (Alpha < alphaLim1) Or (Alpha > alphaLim2)

If Erreur Then

  If Sgn(RLim1) = Sgn(RLim2) Or RLim1 = 0 Or RLim2 = 0 Then
    If Sgn(R) <> Sgn(RLim1) Then
      Mess1 IDm_ChangerSigneRayon
    Else
      Mess1 IDm_RayonCompris & AffGrille(minmaj(RLim1, 1)) & ID_Et & AffGrille(minmaj(RLim2, -1))
    End If
  Else  ' RLim1 et RLim2 de signes différents : le rayon mini = R(-alphadeb)
    Mess2 IDm_Augmenter & ID_AmplitudeRayon, ID_Minimum, "", Max(Abs(RLim1), Abs(RLim2)), 1, IDm_OuLeChangerDeSigne
  End If

End If
  
      
End Sub

Private Function L_a(ByVal param As String, Optional ByVal ChampEnCours As Variant, Optional ByVal Dernier As Variant) As Variant
' calcul de la vitesse de braquage en fonction de la longueur ou inversement
' Si param="longueur", ChampEnCours représente le champ en cours de saisie (et de validation) : rayon, angle de giration ou longueur
' Si ChampEnCours=longueur, Dernier représente  le champ précédemment saisi (paramétrage du message) : rayon ou angle
Dim w As String

        If param = "longueur" Then
          L_a = (v / l) * (Alpha - Alphadeb)
          
              ' Controle des bornes : (aVehMin=0.0005) < a < aVehMax
              '-----------------------------------------------------------------------------
          If Abs(L_a) > aVehMax Then
            Select Case ChampEnCours
            
            Case RAY        ' rayon                               ' Réf. Rectificatif du 24/04/97    - p2 - cas L|R
              If Sgn(Alpha) = Sgn(Alphadeb) Then
                Mess1 IDm_PrendreRayonPlusProche
              ElseIf Alphadeb = 0 Then
                Mess1 IDm_Augmenter & ID_AmplitudeRayon
              ElseIf Abs(v * Alphadeb / l) < aVehMax Then
                Mess1 IDm_Augmenter & ID_AmplitudeRayon & vbCrLf & IDm_OuLeChangerDeSigne
              Else
                Mess1 IDm_ChangerSigneRayon
              End If
              
            Case ANGGIR     ' angle de giration
                              ' devenu inutile suite au document RAC_AL.DOC du 17/06/97 :  géré par CalcAlpha (param="" : cas deltabeta|L)
             
            Case Else '  longueur
              If Dernier = RAY Then        ' Réf. Rectificatif du 24/04/97    - p2 - cas R|L
                Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Abs(v / aVehMax * (Alpha - Alphadeb)), 1
              Else        ' dernier = Angle de giration
                              ' devenu inutile suite au document RAC_AL.DOC du 17/06/97 :  géré par CalcAlpha (param="" :  cas L|deltabeta)
              End If
            End Select
            
            
          ElseIf Abs(angConv(L_a, unite)) < aVehMin Then
            Mess1 IDm_VitBraqTropFaibleInduite
          End If

          Grille.Col = 5
          Grille.Value = AffGrille(angConv(L_a, unite))
        
        Else    ' Calcul de L
          L_a = (v / a) * (Alpha - Alphadeb)
          If L_a / pas_roue > nbIteration Then
            Grille.Col = 5      '        ' Car la vitesse de braquage a pu être déjà recalculée et affichée
            w = Grille.Value
            retrouveLigne 5
            MessConfirm IDm_LgDepltImportant & vbCrLf & IDm_Confirmation, vbYesNo
            Grille.Value = w
          End If
          Grille.Col = 4
          Grille.Value = AffGrille(Abs(L_a))
        End If
  
End Function

Private Function calcAlpha(ByVal param As String, Optional ByVal ChampEnCours As Variant) As Variant
' calcul de alpha dans un RACCORD en fonction de la longueur ou de l'angle de giration ou des 2
'param=     "longueur" pour un calcul à partir de L (a est également connu)                   cas L|a et a|L
'                 "anggir"      pour un calcul à partir de deltabeta (a est également connu)      cas deltabeta|a et a|deltabeta
'                 ""                pour un calcul à partir de  deltabeta et L                                             cas deltabeta|L et L|deltabeta
' ChampEnCours désigne le champ en cours de validation
  
Dim wAlpha As Double, alphaMini As Double, alphaMaxi As Double, deltabetamini As Double, deltabetamaxi As Double
Dim pTmp As Double, pMini1 As Double, pMini2 As Double, pMaxi1 As Double, pMaxi2 As Double
Dim pAlpha As Double, pMini As Double, pMaxi As Double, mini As Double, maxi As Double, dalpha As Double
Dim Lmin As Double, Lmin2 As Double, Lmax As Double
Dim Msg As String, Msg0 As String
Dim libelle As String
Dim Valeur As Double, valeur0 As Double, valeurMaxi As Double
Dim Erreur As Boolean
  
  Select Case param
  Case "longueur"
  ' Référence Complément au Dossier de programmation 14/02/97 (erreur.doc) - Messages d'erreur des raccords p3(bas) et 4
    calcAlpha = Alphadeb + l * a / v
    If Abs(calcAlpha) > alphaVehMax Then
      If ChampEnCours = LG Then         ' longueur
        valeurMaxi = v / a * (Sgn(a) * alphaVehMax - Alphadeb)
        Msg0 = ID_LaLongueur
      Else    ' ChampEnCours = vitesse de braquage
        valeurMaxi = angConv((v / l) * (Sgn(a) * alphaVehMax - Alphadeb), unite)
        Msg0 = ID_LaVitBraq
      End If
      Mess2 IDm_Reduire & " " & Msg0, ID_Maximum, "", Abs(valeurMaxi), -1
    End If
    
  Case "anggir"
  ' Référence Complément au Dossier de programmation 14/02/97 (erreur.doc) - Messages d'erreur des raccords p2(bas) et 3
    ' Cas a|Deltabeta ou DeltaBeta|a
    
    Select Case Sgn(DeltaBeta / v)
    Case 0    ' deltabeta=0
      If Sgn(a) = Sgn(Alphadeb) Then
        If ChampEnCours = ANGGIR Then     ' angle de giration                    ' Réf. RAC_VB_A.doc du 3/07/97 (cas a|Deltabeta)
          Mess1 IDm_AngleNonNul
        Else
          libelle = ID_LaVitBraq & " doit être "
          If Alphadeb < 0 Then
            libelle = libelle & ID_PositifVitBraq
          Else
            libelle = libelle & ID_NegatifVitBraq
          End If
          Mess1 libelle
        End If
        
      Else
        calcAlpha = -Alphadeb
      End If
      
      Exit Function
      
    Case 1      ' deltabeta/v > 0
      Erreur = (Alphadeb <= 0 And a < 0) ' <= au lieu de < :  voir FAX rectificatif du 17/02/97
      
    Case -1     ' deltabeta/v < 0
      Erreur = (Alphadeb >= 0 And a > 0)    ' >= au lieu de > :  voir FAX rectificatif du 17/02/97
    End Select
    
    If ChampEnCours = ANGGIR Then   ' angle de giration
      Valeur = DeltaBeta
      valeur0 = a
      Msg0 = ID_LAngle
    Else
      Valeur = a
      valeur0 = DeltaBeta
      Msg0 = ID_LaVitBraq
    End If
    
    If Erreur Then
      MessConstr IDm_ChangerLeSigneDe & Msg0
    End If
    
   
    If DeltaBeta * a / v > 0 Then                         ' alpha se rapproche de alphaVehMax (en v.a.)
        valeurMaxi = (Cos(Alphadeb) - Cos(alphaVehMax)) / e(0) * (v / valeur0)
    Else                                                           ' ' alpha se rapproche de 0
        valeurMaxi = (Cos(Alphadeb) - 1) / e(0) * (v / valeur0)
    End If
    
    If Abs(Valeur) > Abs(valeurMaxi) Then
      If valeurMaxi = 0 Then  '     le alpha précédent était déjà en limite (0 ou alphaVehMax)
        Mess1 IDm_ChangerLeSigneDe & Msg0
      Else
        Mess2 IDm_Reduire & " " & Msg0, ID_Maximum, "", Abs(angConv(valeurMaxi, unite)), -1
      End If
    End If
      
    wAlpha = Arccos(Cos(Alphadeb) - DeltaBeta * a * e(0) / v)
    If DeltaBeta * a / v > 0 Then                         ' alpha se rapproche de alphaVehMax (en v.a.)
      calcAlpha = Sgn(DeltaBeta / v) * Abs(wAlpha)
    Else                                                           ' ' alpha se rapproche de 0
      calcAlpha = Sgn(Alphadeb) * Abs(wAlpha)
    End If
      '---------------------- Fin de la référence au Complément du Dossier de programmation
 
  Case Else       ' param = ""
  ' Deltabeta et L permettent de calculer alpha pour un Raccord
  ' Algorithme décrit dans le Complément au Dossier de programmation
  ' du 13 février 97(RACCORD.DOC), complété par le rectificatif du 24/04/97 pour les messages
  
    
    pAlpha = -e(0) * DeltaBeta / l    'Pente du segment [alpha,alphaD]
    
    If ChampEnCours = ANGGIR Then    ' angle de giration    ' Réf. RAC_AL.DOC du 17/06/97
      alphaMaxi = Min(alphaVehMax, Alphadeb + aVehMax * l / v)
      alphaMini = Max(-alphaVehMax, Alphadeb - aVehMax * l / v)
      pMini = pente(alphaMaxi)
      pMaxi = pente(alphaMini)
      
      If Sgn(v) = 1 Then
          deltabetamini = l * (-pMaxi) / e(0)
          deltabetamaxi = l * (-pMini) / e(0)
      Else
          deltabetamini = l * (-pMini) / e(0)
          deltabetamaxi = l * (-pMaxi) / e(0)
      End If
      Msg = IDm_AngleCompris & _
                Format(minmaj(angConv(deltabetamini, unite), 1), FmtGrille) & ID_Et & _
                Format(minmaj(angConv(deltabetamaxi, unite), -1), FmtGrille)
      
      Select Case pAlpha
      Case Is < pMini
        Mess1 Msg
      
      Case Is > pMaxi
        Mess1 Msg
      
      Case pMini
        wAlpha = alphaMaxi
      
      Case pMaxi
        wAlpha = alphaMini
      
      Case 0
        wAlpha = -Alphadeb
      
      Case Else   ' Recherche dichotomique sur [-alphaVehMax,-alphaD] ou [-alphaD,alphaVehMax]
        mini = alphaMini
        maxi = alphaMaxi
    
        If pAlpha > 0 Then
          wAlpha = mini
        Else
          wAlpha = maxi
        End If
        wAlpha = (maxi + mini) / 2
      
          dalpha = maxi - mini
          While Abs(dalpha) > 0.000001
            If pAlpha < pente(wAlpha) Then mini = wAlpha Else maxi = wAlpha
            wAlpha = (mini + maxi) / 2
            dalpha = maxi - mini
          Wend
   
      End Select
  
    Else   ' ChampEnCours = longueur
      Select Case Sgn(DeltaBeta / l)
      Case 1
                  ' Vérification des bornes pour les pentes négatives
        If (Cos(Alphadeb) - (aVehMax) * (DeltaBeta / v) * e(0) > Cos(alphaVehMax)) Then
          pMini = pente(Arccos(Cos(Alphadeb) - (aVehMax) * (DeltaBeta / v) * e(0)))
          ' Attention : l’angle est compris entre 0 et PI/2 - fonctionnement Arcos()
        Else
          pMini = pente(alphaVehMax)
        End If
        pMaxi1 = 0
        pMaxi2 = 0
        If Alphadeb > 0 And (Cos(Alphadeb) - (-aVehMax) * (DeltaBeta / v) * e(0) < 1) Then
          pMaxi1 = pente(Arccos(Cos(Alphadeb) - (-aVehMax) * (DeltaBeta / v) * e(0)))
          pMaxi2 = pente(-Arccos(Cos(Alphadeb) - (-aVehMax) * (DeltaBeta / v) * e(0)))
        End If
        If pMaxi1 > pMaxi2 Then pTmp = pMaxi1: pMaxi1 = pMaxi2:  pMaxi2 = pTmp
        Erreur = (pAlpha < pMini Or pAlpha > pMaxi1) And (pAlpha < pMaxi2 Or pAlpha > 0)
        If Erreur Then
          Lmin = Abs(e(0) * DeltaBeta / pMini)
          If pMaxi1 = pMaxi2 Then
            Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Lmin, 1
          Else
            Lmin2 = Abs(e(0) * DeltaBeta / pMaxi2)
            Lmax = Abs(e(0) * DeltaBeta / pMaxi1)
            Msg = ID_Ou & IDm_Compris & Format(minmaj(Lmin, 1), FmtGrille) & ID_Et & Format(minmaj(Lmax, -1), FmtGrille)
            Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Lmin2, 1, Msg
          End If
        End If

            
          ' Ordonner les bornes Pmaxi1 <= Pmaxi2 <= 0 ‘pour plus de surété sur Arcos()
          ' P doit vérifier :  Pmini <= P <= Pmaxi1  ou Pmaxi2 <= P <= 0
          'Msg si Pmaxi1 = Pmaxi2, Lmin = E*Deltabeta/Pmini
          'sinon L doit être supérieure à E*Deltabeta/Pmaxi2 ou compris entre E*Db/Pmini et E*Db/Pmaxi1Note : toutes termes en valeur absolue.}
      Case -1
          'Vérification des bornes pour les pentes positives
        If (Cos(Alphadeb) - (-aVehMax) * (DeltaBeta / v) * e(0) > Cos(alphaVehMax)) Then
          pMaxi = pente(-Arccos(Cos(Alphadeb) - (-aVehMax) * (DeltaBeta / v) * e(0)))
          ' Attention : l’angle est compris entre 0 et PI/2 - fonctionnement Arcos()
        Else
          pMaxi = pente(-alphaVehMax)
        End If
        pMini1 = 0
        pMini2 = 0
        If Alphadeb < 0 And (Cos(Alphadeb) - (aVehMax) * (DeltaBeta / v) * e(0) < 1) Then
          pMini1 = pente(Arccos(Cos(Alphadeb) - (aVehMax) * (DeltaBeta / v) * e(0)))
          pMini2 = pente(-Arccos(Cos(Alphadeb) - (aVehMax) * (DeltaBeta / v) * e(0)))
        End If
        If pMini1 > pMini2 Then pTmp = pMini1: pMini1 = pMini2:  pMini2 = pTmp
        Erreur = (pAlpha > pMaxi Or pAlpha < pMini2) And (pAlpha > pMini1 Or pAlpha < 0)
        If Erreur Then
          Lmin = Abs(e(0) * DeltaBeta / pMaxi)
          If pMini1 = pMini2 Then
            Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Lmin, 1
          Else
            Lmin2 = Abs(e(0) * DeltaBeta / pMini1)
            Lmax = Abs(e(0) * DeltaBeta / pMini2)
'            Msg = " ou donner une valeur comprise entre " & Format(minmaj(Lmin, 1), FmtGrille) & ID_Et & Format(minmaj(Lmax, -1), FmtGrille)
            Msg = ID_Ou & IDm_Compris & Format(minmaj(Lmin, 1), FmtGrille) & ID_Et & Format(minmaj(Lmax, -1), FmtGrille)
            Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Lmin2, 1, Msg
          End If
        End If
      
      Case Else
        ' Vérification pour les pentes nulles
        Lmin = Abs(2 * Alphadeb / aVehMax * v)
        If Abs(l) < Lmin Then Mess2 IDm_AugmenterLongueur, ID_Minimum, "", Lmin, 1
      End Select
      
        
        pMini = pente(alphaVehMax)
        pMaxi = pente(-alphaVehMax)

        Select Case pAlpha
        Case pMini
          wAlpha = alphaVehMax
      
        Case pMaxi
          wAlpha = -alphaVehMax
      
        Case 0
          wAlpha = -Alphadeb
      
        Case Else   ' Recherche dichotomique sur [-alphaVehMax,-alphaD] ou [-alphaD,alphaVehMax]
    
          If pAlpha > 0 Then
            mini = -alphaVehMax
            maxi = -Alphadeb
            wAlpha = mini
          Else
            mini = -Alphadeb
            maxi = alphaVehMax
            wAlpha = maxi
          End If
          wAlpha = (maxi + mini) / 2
      
            dalpha = maxi - mini
            While Abs(dalpha) > 0.000001
              If pAlpha < pente(wAlpha) Then mini = wAlpha Else maxi = wAlpha
              wAlpha = (mini + maxi) / 2
              dalpha = maxi - mini
            Wend
      
        End Select
        
        
    End If
    
    calcAlpha = wAlpha
  End Select    ' Case param
      
      ' Controles terminaux
    If calcAlpha <> 0 Then
      If Abs(e(0) / Atn(calcAlpha) - Sgn(calcAlpha) * voar(0)) >= INFINI Then calcAlpha = 0
    
    If Abs(calcAlpha - Alphadeb) < angConv(0.001 * l / v, radian) Then
        ' réf. Rectificatif sur les raccords du 24/04/97
      MessConstr IDm_RayonFinalProcheDepart
    End If
  End If
    
End Function

Private Function pente(ByVal Angle As Double) As Double
' Fonction appelée par calcAlpha dans le cas d'un raccord calculé
' avec deltabeta et L (param="" dans calcAlpha)

  If Angle = Alphadeb Then
    pente = -Sin(Alphadeb)
  Else
    pente = (Cos(Angle) - Cos(Alphadeb)) / (Angle - Alphadeb)
  End If
End Function


Private Function SommeDeltaBeta(ByVal Angle As Double, Deltabeta1 As Double, L1 As Double, Deltabeta2 As Double, L2 As Double) As Double
  
' Somme des angles de giration d'un virage pour un angle de roues 'angle' donné
' Réf. Dossier de programmation p5  VIRAGE §2

Dim k1, k2 As Double

  k1 = Angle / a
  k2 = e(0) * a
  L1 = v * k1
  L2 = l - 2 * v * k1
  Deltabeta1 = v / k2 * (1 - Cos(Angle))
  Deltabeta2 = L2 / e(0) * Sin(Angle)
  SommeDeltaBeta = 2 * Deltabeta1 + Deltabeta2

End Function

Private Sub verifDeltabetaL()
Dim kMaxi As Double
'Réf. FAX du 14/05/97 pour vérifier que deltabeta et L (cas n°3) peuvent bien faire un virage
' le k à trouver est 2*v/a
' le kmaxi correspond à aVehMax
' Dans la formule fCourante(alpha) du bas de la page 1, on a mis des valeurs absolues pour respecter l'hypothèse de départ
Dim Erreur As Boolean

    kMaxi = (e(0) * Abs(DeltaBeta) - Abs(l) * Sin(alphaVehMax)) / ((1 - Cos(alphaVehMax)) - alphaVehMax * Sin(alphaVehMax))
    Erreur = kMaxi < 2 * Abs(v) / aVehMax
    If Erreur Then Mess1 IDm_ConstructionImpossibleTouteVitBraq

End Sub

Private Sub verifSigne_a_dAlpha(ChampEnCours As Integer)
' Vérifie que la variation de Alpha est du même signe que a
Dim Rdeb As Double

If Alphadeb <> 0 Then Rdeb = e(0) / Tan(Alphadeb) - Sgn(Alphadeb) * voar(0)

        If Sgn(Alpha - Alphadeb) <> Sgn(a) Then     'gère les scénarios [a R] et [R a]
          If ChampEnCours = RAY Then ' Rayon
            If Sgn(Alphadeb) = Sgn(a) Then                  ' réf. : Rectificatif du 24/04/97  - p2 cas a|R complété par COMPLCDC.DOC du 10/06/97
                If Sgn(Alpha) <> Sgn(a) Then
                  Mess1 IDm_ChangerSigneRayon
                Else
                  Mess2 IDm_Reduire & ID_AmplitudeRayon, ID_Strict, ID_Inferieur, "", Abs(Rdeb), -1
                End If
            ElseIf Alphadeb = 0 Then
              Mess1 IDm_ChangerSigneRayon
            Else
              Mess2 IDm_Augmenter & ID_AmplitudeRayon, ID_Strict, ID_Superieur, Abs(Rdeb), 1, IDm_OuLeChangerDeSigne
            End If
            
          Else        ' ChampEnCours = vitesse de braquage
             MessConstr IDm_ChangerSigneVitbraq
          End If
           
          
      End If

End Sub

Private Sub verifSigne_a_Deltabeta(Optional ChampEnCours As Variant)
Dim reponse As Integer

        If Sgn(DeltaBeta * a) <> Sgn(v) Then
          If Sgn(v) = 1 Then
            Msg = ID_LAngle & ID_Et & ID_LaVitBraq & vbCrLf & IDm_MemeSigne
          Else
            Msg = ID_LAngle & ID_Et & ID_LaVitBraq & vbCrLf & IDm_SigneContraire
          End If
          If IsMissing(ChampEnCours) Then
            Mess1 Msg
          Else
            MessConfirm Msg & Chr(13) & Chr(13) & IDm_RedefinirAutreValeur, vbYesNo
'            reponse = MsgBox(Msg & Chr(13) & Chr(13) & "Redéfinir l'autre valeur ?", vbYesNo + vbQuestion + vbDefaultButton2)
'            If reponse = vbYes Then
              With TabLigneTraj(numEnreg)
                .Complet = False
                If Not .dessinEffacé Then effSegmentS numEnreg
                If ExportOK And numEnreg = 1 Then ExportOK = False: GriserMenus Etat:=True
                .DernierChampSaisi(0) = .DernierChampSaisi(1)
                .DernierChampSaisi(1) = ChampEnCours
                If ChampEnCours = 4 Then
                  Grille.Col = 3
                  .DeltaBeta = Empty
                Else
                  Grille.Col = 5
                  .Vitbraq = Empty
                End If
                Grille.Value = ""
                Grille.Action = 0   ' rendre courante la cellule que l'utlisateur vient d'accepter de ressaisir
              End With
            Grille.SetFocus
            Err.Raise 111
          ' Incompatibilité des signes de a et DeltaBeta
          End If
      End If

End Sub

'*********************************************************************************
' Suppression d'un déplacement
'*********************************************************************************
Public Sub SupprDepl()
Dim i%, numdepl%
' L'indicateur Suppr permet de savoir si on peut faire recoller la ligne au-dessus et celle au-dessous (sinon, garder 1 ligne blanche)
Dim Suppr As Boolean
Dim Cancel As Boolean
Dim Msg As String

Dim AlphaFin As Double

 verifSelection Cancel
 If Cancel Then Exit Sub
  
If ligneSel = 1 Then
    Msg = IDm_SupprDebutInterdit: Cancel = True
ElseIf ligneSel = Grille.MaxRows Then
    Msg = IDm_SupprFinInterdit: Cancel = True
ElseIf MsgBox(IDm_SupprIrreversible, vbOKCancel + vbExclamation + vbDefaultButton2, IDm_ConfirmSuppr) = vbCancel Then
    Cancel = True
End If

If Cancel Then
    If Msg <> "" Then MessA Msg Else Grille.SetFocus
    Grille.Action = 14                  ' désactivation de la ligne
    Exit Sub
End If


'------------- Réalisation effective de la suppression ----------------------
Screen.MousePointer = vbHourglass ' Sablier
numEnreg = ligneSel - 1

Modif = True
If FichierJournal Then Write #numFichLog, "Suppression de la ligne " & CStr(ligneSel)

  If TabLigneTraj(numEnreg - 1).Complet Then
    ' Déterminer si on peut raccrocher les 2 lignes (Supppr=True) ou non
    With TabLigneTraj(numEnreg + 1)
      ' Détermine si on peut ressouder les 2 lignes suite à la suppression  de la ligne ou si on doit laisser une ligne blanche
      If .deplacement = LIGNE And TabLigneTraj(numEnreg - 1).deplacement = LIGNE Then  ' pas 2 lignes consécutives
'      ElseIf .rayon = Empty And Not .Complet Then
      ElseIf IsEmpty(.Rayon) Then      ' Debuf : modif AV (18/09/97 : si  rayon Empty, Déplacement forcément incomplet)
        Suppr = True
      Else
      ' Recherche de raccordement des rayons (Alphadeb =   AlphaFin)
        If .deplacement = VIRAGE Then
          Alphadeb = 0
        Else
          Alphadeb = .Alpha - .Longueur * .Vitbraq / v    ' Inversion de la formule du Dossier de progr.  p.6 "RACCORD" (R et deltabeta)
        End If
        With TabLigneTraj(numEnreg - 1)
          If .deplacement = VIRAGE Then
            AlphaFin = 0
          Else
            AlphaFin = .Alpha
          End If
        End With
  
        Suppr = Abs(Alphadeb - AlphaFin) < 0.0001
      End If
    End With
  ElseIf TabLigneTraj(numEnreg - 1).deplacement = -1 Then
  ' Ce cas particulier, fait à la demande du CERTU (10/10/97) permet d'éviter d'avoir 2 lignes blanches consécutives
  ' il faut éviter de redessiner les segments suivants (voir + loin)
    Suppr = True
  End If
  
  For i = numEnreg To Grille.MaxRows - 2
    effSegment TabLigneTraj(i)
  Next

  If Suppr Then
    With Grille
      .Row = numEnreg + 1
      .Action = 5     ' suppression
      .MaxRows = .MaxRows - 1
      For i = numEnreg + 1 To .MaxRows
        TabLigneTraj(i - 1) = TabLigneTraj(i)
        activeCol i
        If TabLigneTraj(i).Complet And TabLigneTraj(i).dessinEffacé And TabLigneTraj(numEnreg - 1).deplacement <> -1 Then desSegment TabLigneTraj(i - 1), TabLigneTraj(i - 2)
        numdepl = TabLigneTraj(i - 2).deplacement
        
        ' Réautorisation des champs verrouillés suite à insertion
        Grille.Row = i
        If TabLigneTraj(i - 2).Complet Then autorChamp
        Grille.Row = i - 1
        If (numdepl = RACCORD Or numdepl = DEBUT) And Not TabLigneTraj(i).Complet Then
          Grille.Col = 2
          Grille.Lock = False
        End If
      Next
      ReDim Preserve TabLigneTraj(.MaxRows + 1)
      activeCol numEnreg + 1
    End With
  
  Else
    numdepl = TabLigneTraj(numEnreg - 1).deplacement
    If (numdepl = RACCORD Or numdepl = DEBUT) Then
      Grille.Row = numEnreg
      Grille.Col = 2
      Grille.Lock = False
    End If
    reinLigne numEnreg
    Grille.TypeComboBoxCurSel = -1
  End If
  
  VerifExport
  
  calculPositions
  
  chgEnreg numEnreg + 1
  Grille.SetFocus
  Screen.MousePointer = vbDefault

End Sub

'*********************************************************************************
' Insertion d'un déplacement
'*********************************************************************************
Public Sub InserDepl(Optional ByVal DecomposVirage As Variant)
' DecomposVirage indique que  l'appel a été fait depuis la fonction decompVirage
Dim i%, j%, memoNumEnreg%
Dim Cancel As Boolean
Dim Msg As String

If IsMissing(DecomposVirage) Then
  
  verifSelection Cancel
  If Cancel Then Exit Sub
  
  If ligneSel = 1 Then
    Msg = IDm_InsertionAudebutInterdit: Cancel = True
  ElseIf VersionDemo And Grille.MaxRows > MAXLIGNEDEMO Then
    Msg = ID_ExemplDemo & vbCrLf & IDm_LimiteAtteinte: Cancel = True
  End If

  If Cancel Then
    If Msg <> "" Then MessA Msg
    Grille.Action = 14                  ' désactivation de la ligne
    Exit Sub
  End If
  
End If    ' Not DecomposVirage

'------------- Réalisation effective de l'Insertion ----------------------
Modif = True
If IsMissing(DecomposVirage) Then If FichierJournal Then Write #numFichLog, "Insertion à la ligne " & CStr(ligneSel)

Screen.MousePointer = vbHourglass ' Sablier
memoNumEnreg = numEnreg
numEnreg = ligneSel - 1

  With Grille
    .Row = numEnreg + 1
    .Action = 7       ' insertion
    ReDim Preserve TabLigneTraj(.MaxRows)
    For i = .MaxRows - 1 To numEnreg + 1 Step -1
    ' Transfert de chaque déplacement sur la ligne suivante
      TabLigneTraj(i) = TabLigneTraj(i - 1)
      activeCol i + 1
      If IsMissing(DecomposVirage) Then
        effSegment TabLigneTraj(i)
        For j = 2 To 5: .Col = j: .Lock = True: Next
      End If
    Next
    ' Suite à l'affectation ci-dessus, le point de la ligne créée pointe sur celui de la ligne qui suit
   Set TabLigneTraj(i).p = Nothing
    
    numEnreg = ligneSel - 1
'    With TabLigneTraj(memoNumEnreg - 1)
    With TabLigneTraj(numEnreg - 1)
      If IsMissing(DecomposVirage) And (.deplacement = RACCORD Or .deplacement = DEBUT) Then
      ' Réautorisation du Rayon précédent l'insertion, la ligne étant redevenue vierge
'        activeCol memoNumEnreg
        activeCol ligneSel - 1
        Grille.Col = 2
        Grille.Lock = False
      End If
    End With
      
    reinLigne numEnreg
    creLigne
'    NumEnreg = memoNumEnreg
  End With

calculPositions

Screen.MousePointer = vbDefault

End Sub

'*********************************************************************************
' Décomposition d'un virage
'*********************************************************************************
Public Sub DECOMPVIRAGE()
' Décomposition de virage
Dim i%, j%, maxLS%
Dim CercleNul As Boolean
Dim Cancel As Boolean
Dim Msg As String

  verifSelection Cancel
  If Cancel Then Exit Sub
  
  numEnreg = ligneSel - 1
  
  If TabLigneTraj(numEnreg).deplacement <> VIRAGE Then
    Msg = IDm_SelectVirage: Cancel = True
  ElseIf Not TabLigneTraj(numEnreg).Complet Then
    Msg = IDm_VirageIncomplet: Cancel = True
  ElseIf VersionDemo And Grille.MaxRows > MAXLIGNEDEMO - 1 Then
    Msg = ID_ExemplDemo & vbCrLf & IDm_LimiteAtteinte: Cancel = True
  End If

  If Cancel Then
    MessA Msg
    Grille.Action = 14  ' Désactivation de la sélection
    Exit Sub
  End If

Grille.Action = 14                  ' désactivation de la ligne
Grille.OperationMode = 2      ' Mode Ligne en surbrillance
activeCol ligneSel
Screen.MousePointer = vbHourglass ' Sablier

If MsgBox(IDm_DecomposIrreversible, vbOKCancel + vbExclamation + vbDefaultButton2, IDm_ConfirmDecompos) = vbOK Then
'------------- Réalisation effective de la décomposition ----------------------
  Modif = True
  If FichierJournal Then Write #numFichLog, "Décomposition de la ligne " & CStr(ligneSel)
  
  CercleNul = (Abs(TabVirage(TabLigneTraj(numEnreg).numVirage).element(1).Longueur) < 0.001)
  For i = 0 To 1
    numEnreg = numEnreg + 1
    InserDepl True
    If CercleNul Then Exit For  ' on a besoin d'1 ou 2 insertions selon que l'on a un cercle ou non
  Next

'  NumEnreg = NumEnreg - 2
If CercleNul Then
  maxLS = ligneSel
Else
  maxLS = ligneSel + 1
End If
' NumEnreg = maxim
j = maxLS
  
  For i = 2 To 0 Step -1
    If CercleNul And i = 1 Then
    Else
      TabLigneTraj(j) = TabVirage(TabLigneTraj(numEnreg).numVirage).element(i)
      With TabLigneTraj(j)
        .Complet = True
        If .deplacement = RACCORD Then
          .DernierChampSaisi(1) = 1         ' Rayon
          .DernierChampSaisi(0) = 2        ' Angle (pour mettre quelque chose)
        End If
      End With
      alimGrille j, TabLigneTraj(j)
      
      Grille.Col = 2      ' bloquage ou non du rayon
      If TabLigneTraj(j + 1).dessinCalculé Then
        Grille.Lock = True
      End If
      j = j - 1
    End If
  Next

  Grille.Action = 0 ' pour faire coïncider numenreg avec la ligne en cours
  Grille.EditEnterAction = 5
End If
  
VerifExport

' Pour que le Rafraichir suivant recalcule tout le dessin
flagRecalcul = True
AfficheMarque TabLigneTraj(numEnreg), 0

Grille.OperationMode = 0
Grille.SetFocus
Screen.MousePointer = vbDefault

End Sub


Private Sub verifSelection(Cancel As Boolean)

  If Not Grille.IsBlockSelected Or ligneSel = -1 Then
    MessA IDm_SelectLigne
    Cancel = True
    Exit Sub
  ElseIf Grille.SelBlockRow <> Grille.SelBlockRow2 Then   'Plusieurs lignes sélectionnées.
    MessA IDm_SelectUneLigne
    Grille.Action = 14                  ' désactivation de la sélection
    Cancel = True
    Exit Sub
  End If
  
End Sub


Private Sub VerifExport()
Dim i%

  For i = 1 To Grille.MaxRows - 1
    If Not TabLigneTraj(i - 1).Complet Then Exit For
  Next
  ExportOK = (i = Grille.MaxRows)
  GriserMenus Etat:=True

End Sub

'*********************************************************************************************
' Ecriture des caractéristiques de la trajectoire dans un fichier au format texte
'*********************************************************************************************
Public Sub ecrireFormatTexte()
Dim i%, j%, numFich%, num%(4)
Dim fv As Form

On Error GoTo ErrHandler

  With dlgExportTexte
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    If .InitDir = "" Then .InitDir = App.Path       ' pas de répertoire par défaut prévu  : répertoire courant
    .ShowSave
    If FichierProtégé(.FileName) Then Exit Sub
    On Error GoTo 0
  End With
  
  ActivMenu Activation:=False

  numFich = FreeFile
  Open dlgExportTexte.FileName For Output Lock Read Write As numFich
      If FichierJournal Then Write #numFichLog, "Enregistrer Caractéristiques : " & dlgExportTexte.FileName
  
  Print #numFich, "GIRATION " & GirationVersion; vbTab;
  Print #numFich, ID_MaitriseOuvrage; vbCrLf
  
  Load frmVéhicule
    ' m/s-->Km/h
  With frmVéhicule
'    Print #numFich, Majus(ID_Véhicule) & vbTab; Trim(.Nom) & vbTab; "Vitesse" & vbTab; v * 3.6 & vbCrLf
'    Print #numFich, Majus(ID_Véhicule) & vbTab; lblCaracVeh & vbCrLf
    EcrireLigne numFich, Majus(ID_Véhicule), Trim(.lblNom)
    
    Select Case Vehic.nbRemorque
    Case 0
      Print #numFich, vbCrLf
      EcrireLigne numFich, ElimEnTete(.lblPAFav), .txtPAFav(0)
      EcrireLigne numFich, ElimEnTete(.lblE), .txtE(0)
      EcrireLigne numFich, ElimEnTete(.lblPAFar), .txtPAFar(0)
      EcrireLigne numFich, ElimEnTete(.lblL), .txtL(0)
      EcrireLigne numFich, ElimEnTete(.lblLEav), .txtLEav
      EcrireLigne numFich, ElimEnTete(.lblLEar), .txtLEar(0) ',vbcrlf : 2/5/2000 (à rajouter si c'est demandé)
    Case 1
      Print #numFich, vbTab;
      EcrireLigne numFich, .fraTracteur, .fraRemorque1
      EcrireLigne numFich, ElimEnTete(.lblPAFav), .txtPAFav(0), .txtPAFav(1)
      EcrireLigne numFich, ElimEnTete(.lblE), .txtE(0), .txtE(1)
      EcrireLigne numFich, ElimEnTete(.lblPAFar), .txtPAFar(0), .txtPAFar(1)
      EcrireLigne numFich, ElimEnTete(.lblL), .txtL(0), .txtL(1)
      EcrireLigne numFich, ElimEnTete(.lblLEav), .txtLEav
      EcrireLigne numFich, ElimEnTete(.lblLEar), .txtLEar(0), .txtLEar(1)
      EcrireLigne numFich, ElimEnTete(.lblDpat), .txtDpat(0) ',vbcrlf : 2/5/2000 (à rajouter si c'est demandé)
    Case 2
      Print #numFich, vbTab;
      EcrireLigne numFich, .fraTracteur, .fraRemorque1, .fraRemorque2
      EcrireLigne numFich, ElimEnTete(.lblPAFav), .txtPAFav(0), .txtPAFav(1), .txtPAFav(2)
      EcrireLigne numFich, ElimEnTete(.lblE), .txtE(0), .txtE(1), .txtE(2)
      EcrireLigne numFich, ElimEnTete(.lblPAFar), .txtPAFar(0), .txtPAFar(1), .txtPAFar(2)
      EcrireLigne numFich, ElimEnTete(.lblL), .txtL(0), .txtL(1), .txtL(2)
      EcrireLigne numFich, ElimEnTete(.lblLEav), .txtLEav
      EcrireLigne numFich, ElimEnTete(.lblLEar), .txtLEar(0), .txtLEar(1), .txtLEar(2)
      EcrireLigne numFich, ElimEnTete(.lblDpat), .txtDpat(0), .txtDpat(1) ',vbcrlf: 2/5/2000 (à rajouter si c'est demandé)
    End Select
    
    EcrireLigne numFich, .fraCarTrac
    EcrireLigne numFich, .lblAVehMax & " " & .lblUnite, .txtaVehMax
    EcrireLigne numFich, .cboRayon, .txtRTracteur ',vbcrlf: 2/5/2000 (à rajouter si c'est demandé)
    EcrireLigne numFich, .fraCarVeh
    EcrireLigne numFich, "", .lblEntreMurs, .lblEntreTrottoirs
    EcrireLigne numFich, .lblRayInter, .lblRayVehic(1), .lblRayVehic(0)
    EcrireLigne numFich, .lblRayExter, .lblRayVehic(3), .lblRayVehic(2), vbCrLf
  End With
  
  
  Unload frmVéhicule
  
'  EcrireLigne numFich, Majus(ID_Trajectoire), Caption
  With TabLigneTraj(0)   ' élément début  --> origine
    EcrireLigne numFich, ID_Origine, .p.X, .p.Y, angConv(.Beta, Enradian:=False)
  End With
  EcrireLigne numFich, ID_Vitesse, Arrond(v * 3.6, 2) & " " & ID_UniteVitesse
  EcrireLigne numFich, ID_UnitAngle, Left(libUnite(OptGen.unite), Len(libUnite(OptGen.unite)) - 3), vbCrLf
'  EcrireLigne numFich, ID_Deplacement, ID_Rayongir, ID_Anglegir, ID_Longueur, ID_VitBraq
  EcrireLigne numFich, "", Left(ID_Rayongir, 1), Left(ID_Anglegir, 1), Left(ID_Longueur, 1), Left(ID_VitBraq, 1)
  For i = 0 To UBound(TabLigneTraj) - 1
    With TabLigneTraj(i)
      If .Complet Then
        Grille.Row = i + 1
        For j = 1 To 5
          Grille.Col = j
          Print #numFich, Grille.Text; vbTab;
        Next
        Print #numFich, vbCrLf;
      ElseIf TabLigneTraj(i).deplacement <> -1 Then
        MessA (IDm_TrajectoireIncomplet & " " & CStr(i))
      Else
        For j = i + 1 To UBound(TabLigneTraj) - 1 ' on affiche un warning dès qu'on tombe sur une ligne non blanche
          If TabLigneTraj(j).deplacement <> -1 Then
            MessA (IDm_TrajectoireIncomplet & " " & CStr(i))
          End If
        Next
        Exit For
      End If
    End With
  Next
  
  Print #numFich, vbCrLf;
  EcrireLigne numFich, ID_NomsFichiers
  EcrireLigne numFich, Majus(ID_Trajectoire), dlgTrajectoire.FileName
  EcrireLigne numFich, Majus(ID_FDP), FDP
  Close numFich
  
  If FichierJournal Then Write #numFichLog, "Enregistrer terminé"
  
   ActivMenu Activation:=True
   GriserMenus Etat:=True
     
  Exit Sub

ErrHandler:
  If Err = cdlCancel Then Exit Sub Else ErrGeneral
   
End Sub

Private Sub EcrireLigne(ByVal numFich As Integer, ParamArray Texte())
Dim i As Integer
  Print #numFich, Texte(0);
  For i = 1 To UBound(Texte)
    Select Case VarType(Texte(i))
    Case vbSingle, vbDouble
      Texte(i) = Format(Texte(i), "0.0##")
    End Select
    Print #numFich, vbTab; Texte(i);
  Next
  Print #numFich, vbCrLf;
End Sub

Private Function ElimEnTete(label As String) As String

  ElimEnTete = Mid(label, InStr(label, ":") + 2)
  
End Function


'*********************************************************************************************
' Ecriture des caractéristiques de la trajectoire dans un fichier au format Giration (.Gir)
'*********************************************************************************************
Public Function ecrire(ByVal NomFich As String) As Boolean
Dim i As Integer, j As Integer, numFich As Integer, num(4) As Integer

  
' cette notation permet d'être homogène avec la v2
  num(DEBUT) = 1
  num(LIGNE) = 2
  num(CERCLE) = 3
  num(VIRAGE) = 4
  num(RACCORD) = 5
  
  ActivMenu Activation:=False
  
  If numFichProtégé <> 0 Then Close numFichProtégé
  
  numFich = FreeFile
  Open NomFich For Output Lock Read Write As numFich
      If FichierJournal Then Write #numFichLog, "Enregistrer " & dlgTrajectoire.FileName
  
  If Extension(FDP, vbLowerCase) <> "fdp" Then
    Write #numFich, ""
  Else
    Write #numFich, nomCourt(FDP)
  End If
  With Vehic
    Write #numFich, Trim(.Nom)
    Write #numFich, v, .PAFav(0), .e(0), .PAFar(0), .l(0) / 2, .LEav / 2, .LEar(0) / 2, alphaVehMax, .aVehMax
    For i = 1 To .nbRemorque
      Write #numFich, i, .Dpat(i - 1), .PAFav(i), .e(i), .PAFar(i), .l(i) / 2, .LEar(i) / 2
    Next
  End With
        
  With TabLigneTraj(0)   ' élément début  --> origine
    Write #numFich, 0, .p.X, .p.Y, .Beta    ' le 0 sert d'indicateur qu'il n'y a plus de remorque
  End With
  
  For i = 0 To UBound(TabLigneTraj) - 1
    With TabLigneTraj(i)
      If .Complet Then
        Write #numFich, num(.deplacement), .Alpha, .DeltaBeta, Abs(.Longueur), .Vitbraq
      ElseIf TabLigneTraj(i).deplacement <> -1 Then
        MessA (IDm_TrajectoireIncomplet & " " & CStr(i))
      Else
        For j = i + 1 To UBound(TabLigneTraj) - 1 ' on affiche un warning dès qu'on tombe sur une ligne non blanche
          If TabLigneTraj(j).deplacement <> -1 Then
            MessA (IDm_TrajectoireIncomplet & " " & CStr(i))
            Exit For
          End If
        Next
        Exit For
      End If
    End With
  Next
  
  Close numFich
  
  ' Verrouillage du fichier le temps de la session
  numFichProtégé = FreeFile
  Open NomFich For Random Lock Read As numFichProtégé
  
  If FichierJournal Then Write #numFichLog, "Enregistrer terminé"
  
  Modif = False
  ActivMenu Activation:=True
  GriserMenus Etat:=True
   
End Function

Public Function lire() As Boolean
Dim i%, j%, numFich%, numligne%, num%(5)
Dim Deltabeta1 As Double, Deltabeta2 As Double, L1 As Double, L2 As Double
Dim v2 As Boolean, reponse As Integer, NomVehic As String
Dim w As String
Dim X As Double, Y As Double
Dim NomFich As String

ActivMenu Activation:=False

' cette notation permet d'être homogène avec la v2
  num(1) = DEBUT
  num(2) = LIGNE
  num(3) = CERCLE
  num(4) = VIRAGE
  num(5) = RACCORD
  
  NomFich = MDIGiration.dlgTrajectoire.FileName
  v2 = (Extension(NomFich, vbLowerCase) = "egt")
  
  If FichierProtégé(NomFich, , Titre:=NomFich, LectureSeuleAutorisée:=True) Then
    lire = True
    Exit Function
  End If
  
  numFich = FreeFile
  Open NomFich For Input As numFich
  
  With Vehic
    
    If v2 Then
      Input #numFich, NomVehic
      .Nom = Trim(NomVehic)
      Input #numFich, w
      v = decmpEnreg(w)
      PAFav(0) = decmpEnreg(w)
      e(0) = decmpEnreg(w)
      PAFar(0) = decmpEnreg(w)
      larg(0) = decmpEnreg(w)
      voav = decmpEnreg(w)
      voar(0) = decmpEnreg(w)
      alphaVehMax = decmpEnreg(w)
      aVehMax = decmpEnreg(w)
    Else
      Input #numFich, FDP
      If FDP <> "" Then
        If nomCourt(FDP) = FDP Then FDP = OptGen.Repert(0) & "\" & FDP & ".fdp"
        If Not ExistFich(FDP) Then
          reponse = MsgBox(Majus(ID_Fichier) & " " & ID_FDP & " " & FDP & IDm_Absent & vbCrLf & ID_Rechercher, vbYesNoCancel + vbQuestion)
          FDP = ""
          Select Case reponse
          Case vbYes
            Import True   ' true: indique que la lecture effective ne sera faite que plus tard
          Case vbCancel
            lire = True
            Exit Function
          End Select
        End If
      End If

      Input #numFich, NomVehic
      .Nom = Trim(NomVehic)
      Input #numFich, v, PAFav(0), e(0), PAFar(0), larg(0), voav, voar(0), alphaVehMax, aVehMax
    End If
    
    .PAFav(0) = PAFav(0)
    .e(0) = e(0)
    .PAFar(0) = PAFar(0)
    .l(0) = 2 * larg(0)
    .LEav = 2 * voav
    .LEar(0) = 2 * voar(0)
    .aVehMax = aVehMax
     .RVehMin = e(0) / Tan(alphaVehMax) - voar(0)        ' Référence: Dossier de programmation - Bibliothèque de véhicules - §1
     RVehMin = .RVehMin
    
    
    numligne = 0
    ReDim TabLigneTraj(0)
    
    Input #numFich, i
    If v2 Then ' fichier egt
      .nbRemorque = i
      nbRemorque = .nbRemorque
      ReDim TabLigneTraj(0).MemoPt(nbRemorque, 0)
      ReDim TabLigneTraj(0).MemoBeta(nbRemorque, 0)
      If i = 1 Then
        Input #numFich, w
        Dpat(i - 1) = decmpEnreg(w)
        PAFav(i) = decmpEnreg(w)
        e(i) = decmpEnreg(w)
        PAFar(i) = decmpEnreg(w)
        larg(i) = decmpEnreg(w)
        voar(i) = decmpEnreg(w)
        
        .Dpat(0) = Dpat(i - 1)
        .PAFav(i) = PAFav(i)
        .e(i) = e(i)
        .PAFar(i) = PAFar(i)
        .l(i) = 2 * larg(i)
        .LEar(i) = 2 * voar(i)
        transRemorque(0) = -(e(i - 1) + Dpat(i - 1))
      End If
      
      ' origine du tracteur
      txtOri(0) = 0
      txtOri(1) = 0
      txtOri(2) = 0
    
    Else          ' .GIR
      
      For j = 1 To 2     'remorques
        If i = 0 Then
          Exit For ' plus de remorque(s)
        Else
          .nbRemorque = j
          ReDim TabLigneTraj(0).MemoPt(nbRemorque, 0)
          ReDim TabLigneTraj(0).MemoBeta(nbRemorque, 0)
          
          Input #numFich, Dpat(i - 1), PAFav(i), e(i), PAFar(i), larg(i), voar(i)
          transRemorque(i - 1) = -(e(i - 1) + Dpat(i - 1))
          
          .Dpat(i - 1) = Dpat(i - 1)
          .PAFav(i) = PAFav(i)
          .e(i) = e(i)
          .PAFar(i) = PAFar(i)
          .l(i) = 2 * larg(i)
          .LEar(i) = 2 * voar(i)
        End If
        Input #numFich, i
      Next

      nbRemorque = .nbRemorque
      ReDim TabLigneTraj(0).MemoPt(nbRemorque, 0)
      ReDim TabLigneTraj(0).MemoBeta(nbRemorque, 0)

      With TabLigneTraj(0)   ' Origine
        Input #numFich, X, Y, Beta
        .p.X = X
        .p.Y = Y
        .Beta = Beta
        
        IgnoreOrigine = True    ' la fonction txtOri_Change peut être déclenchée intempestivement
        txtOri(0) = Arrond(.p.X, 2)
        txtOri(1) = Arrond(.p.Y, 2)
        txtOri(2) = Arrond(angConv(.Beta, unite), 2)
        IgnoreOrigine = False
        betadeb = .Beta
        .MemoPt(0, 0).X = X
        .MemoPt(0, 0).Y = Y
        .MemoBeta(0, 0) = Beta
        
        For i = 1 To nbRemorque
          .MemoPt(i, 0).X = .MemoPt(i - 1, 0).X + transRemorque(i - 1) * Cos(.Beta)
          .MemoPt(i, 0).Y = .MemoPt(i - 1, 0).Y + transRemorque(i - 1) * Sin(.Beta)
          .MemoBeta(i, 0) = .Beta
          If i = 1 Then
            .Sigma = .Beta
          Else
            .Gamma = .Beta
          End If
        Next
          
      End With
      
    End If    ' V2
      
    Do While Not EOF(numFich)
      With TabLigneTraj(numligne)
        Input #numFich, .deplacement
        .deplacement = num(.deplacement)
        If v2 Then
          Input #numFich, w
          Alpha = decmpEnreg(w)
          If Alpha <> 0 Then    ' les fichiers EGT peuvent contenir des alphas très voisins de  : GIRATION considère le rayon INFINI
            If Abs(e(0) / Tan(Alpha) - Sgn(Alpha) * voar(0)) >= INFINI Then Alpha = 0   ' Dossier de progr. p3 et p,5 § "CERCLE-1"
          End If
          DeltaBeta = decmpEnreg(w)
          l = decmpEnreg(w)
          a = decmpEnreg(w)
        Else
          Input #numFich, Alpha, DeltaBeta, l, a
        End If
        .Alpha = Alpha
        .DeltaBeta = DeltaBeta
        .Longueur = Sgn(v) * l
        l = .Longueur
        .Vitbraq = a

          Grille.Row = numligne + 1
          alimGrille numligne, TabLigneTraj(numligne)
          
            If .deplacement = VIRAGE Then
              SommeDeltaBeta .Alpha, Deltabeta1, L1, Deltabeta2, L2
              .numVirage = nbVirage
              nbVirage = nbVirage + 1
              ReDim Preserve TabVirage(.numVirage)
              With TabVirage(.numVirage)
                With .element(0)
                  .deplacement = RACCORD
                  .Rayon = R
                  .DeltaBeta = Deltabeta1
                  .Longueur = L1
                  .Vitbraq = a
                  .Alpha = Alpha
                  .Beta = betadeb + Deltabeta1
                End With
                With .element(1)
                  .deplacement = CERCLE
                  .Rayon = R
                  .DeltaBeta = Deltabeta2
                  .Longueur = L2
                  .Vitbraq = 0
                  .Alpha = Alpha
                  .Beta = betadeb + Deltabeta1 + Deltabeta2
                End With
                With .element(2)
                  .deplacement = RACCORD
                  .Rayon = 0
                  .DeltaBeta = Deltabeta1
                  .Longueur = L1
                  .Vitbraq = -a
                  .Alpha = 0
                  .Beta = betadeb + .DeltaBeta
                End With
              End With    ' tabVirage(numVirage)
            End If    ' cas du virage
            
      ' On alimente les indicateurs Rayon et Angle de giration pour les raccords
            If .deplacement = RACCORD Then .DernierChampSaisi(1) = 1: .DernierChampSaisi(0) = 2
            
            .Beta = betadeb + .DeltaBeta
            betadeb = .Beta
          .Complet = True
      End With      ' tabLignetraj(numligne)
      
      numligne = numligne + 1
      creLigne
      If VersionDemo And numligne = MAXLIGNEDEMO Then Exit Do
    Loop      ' EOF
    
    Close numFich
  End With
    
  ' Verrouillage du fichier le temps de la session
  numFichProtégé = FreeFile
  Open NomFich For Random Lock Read As numFichProtégé
  
  ExportOK = True
'Grille.Action = 13

End Function

Public Function decmpEnreg(s As String) As Double
' La fonction lire doit relire les fichiers EGT avec des blancs comme séparateurs
' L'instruction Input# ne connait que la virgule comme séparateur et lit alors tout l'enregistrement
' decmpEnreg    décompose l'enregistrement
'                     retourne le 1er champ
'                     décale le pointeur s sur le champ suivant
Dim pos%

  pos = InStr(s, " ")
  If pos = 0 Then
    decmpEnreg = val(s)
  Else
    decmpEnreg = val(Left(s, pos))
    s = Mid(s, pos + 1)
    While (Left(s, 1)) = " "
      s = Mid(s, 2)
    Wend
  End If
  
End Function

Private Sub alimGrille(ByVal numligne As Integer, t As ligneTraj)

    With Grille
      .Row = numligne + 1
      Alpha = t.Alpha    ' rayon
      If t.deplacement = LIGNE Then
        R = 0
      Else
        R = R_alpha("alpha")
      End If
      t.Rayon = R
      .Col = 1      ' type de déplacement
      .Lock = True
      If t.deplacement = DEBUT Then
        Grille.SetText 1, 1, ID_DEBUT
      Else
        .TypeComboBoxCurSel = t.deplacement
        .Col = 2    ' rayon
        If t.deplacement = CERCLE Then
          .Lock = True
          .BackColor = grisPale
        ElseIf t.deplacement = RACCORD Then
          .Lock = False
          .BackColor = blanc
        End If
        .Col = 3    ' angle de giration
        If t.deplacement = LIGNE Then
        Else
          .Value = AffGrille(angConv(t.DeltaBeta, unite))
          .Lock = False
          .BackColor = blanc
        End If
          .Col = 4    ' longueur
          .Value = AffGrille(Abs(t.Longueur))
          .Lock = False
          .BackColor = blanc
          .Col = 5    ' vitesse de braquage
        If t.deplacement = CERCLE Or t.deplacement = LIGNE Then
          .Lock = True
          .BackColor = grisPale
        Else
          .Value = AffGrille(angConv(t.Vitbraq, unite))
          .Lock = False
          .BackColor = blanc
        End If
      End If
        
  End With

End Sub

Private Function AffGrille(ByVal v As Variant) As String
  
  AffGrille = Format(v, FmtGrille)
  
End Function

'***************************************************************************************
' Vérifie qu'un autre zoom ou autre action interactive n'est pas déjà en action
'***************************************************************************************
Private Function VerifZoomPossible(Outil As TOOLZOOMEnum) As Boolean
Dim i As Integer
  
  If Not flagDeuxiemePt Then
    'Sinon : Attente d'un 2ème point (Zoom, PAN, Origine..) : une action est en cours
  
    flagZoom = False
    flagPAN = False
    
    If Not gtInteractionEnCours Then
      For i = TOOL_ZOOM To TOOL_ZTOUT
        MDIGiration.mnuZoom(i).Checked = (i = Outil)
      Next
      gtOutilZoom = Outil
      VerifZoomPossible = True
    End If
  End If

End Function

Private Sub cmdZArriere_Click()
' Zoom arrière

  If VerifZoomPossible(TOOL_ZARRIERE) Then
  
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect

    'Modif v3.3(02/05/07) : On évite ainsi de charger frmCadrage inutilement
    'frmCadrage.cmdZArriere = True
    GirStand.ExecuteZArrière
  End If

End Sub

Public Sub cmdZAvant_Click()
' Zoom avant

  If VerifZoomPossible(TOOL_ZAVANT) Then
  
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect
    'Mise en commentaire v3.3(25/04/07) : personne n'est jamais tombé sur ce message
    'Load frmCadrage
    'If Not frmCadrage.cmdZAvant.Enabled Then MsgBox "Cet appel ne devrait pas arriver", vbExclamation, "Message CTOuest"
    'Modif v3.3(25/04/07) : On évite ainsi de charger frmCadrage inutilement
'    frmCadrage.cmdZAvant = True
    ExecuteZAvant
  End If
  
End Sub

Public Sub cmdZoom_Click()
' Zoom fenêtre
  
  If VerifZoomPossible(TOOL_ZOOM) Then
  
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect
    
    'Modif v3.3(25/04/07) : On évite ainsi de charger frmCadrage inutilement
    'frmCadrage.cmdDésigner(0) = True
    PréparerZoom 0
  End If
    
End Sub

Public Sub cmdPAN_Click()
            
  If VerifZoomPossible(TOOL_PAN) Then

'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect
    
    'Modif v3.3(25/04/07) : On évite ainsi de charger frmCadrage inutilement
    'frmCadrage.cmdDésigner(1) = True
    PréparerZoom 1
  End If

End Sub

Private Sub MessA(ByVal Msg As String)
' Message d'avertissement

  MsgBox Msg, vbExclamation
  If Grille.Visible And Grille.Enabled Then Grille.SetFocus
  
End Sub

Private Sub Mess1(ByVal Msg As String)
' Affichage d'un message d'erreur
  MsgBox Msg, vbExclamation
  If Grille.Visible Then Grille.SetFocus
  Err.Raise 110
  
End Sub

Private Sub Mess2(ByVal Msg As String, ByVal PremierMot As String, ByVal Complement As String, ByVal Valeur As Double, ByVal signe As Integer, Optional ByVal Msg2 As Variant)
' Affichage d'un message d'erreur avec indication de valeur borne
' Le complement vient éventuellement compléter l e mot clé (supérieur, inférieur, minimum, maximum,,,)
' Composition du message : '<message> (<motclé> : <valeur formatée>)'
'
  Valeur = minmaj(Valeur, signe)
  If Complement <> "" Then Complement = " " & Complement
  Msg = Msg & " (" & PremierMot & Complement & " : " & Format(Valeur, FmtGrille) & ")"
  If IsMissing(Msg2) Then
  Else
 '   Rajout d'une 2ème ligne de message
    Msg = Msg & vbCrLf & Msg2
  End If
  
  Mess1 Msg
  Exit Sub
  
  MsgBox Msg, vbExclamation
  
  Err.Raise 110
  
End Sub

Private Sub MessConstr(Msg As String)

  MsgBox Msg, vbExclamation, IDm_ConstructionImpossible
  Err.Raise 110
  
End Sub

Private Sub MessConfirm(ByVal Msg As String, flag As Integer)

  If flag = vbYesNo Then
  ' Demande de confirmation 'lourde' : vbExclamation - on rajoute le ? dans le texte du message
    If MsgBox(Msg & " ?", flag + vbExclamation + vbDefaultButton2) = vbNo Then
      Err.Raise 110
    End If
  Else
  ' Demande de confirmation 'légère' : vbQuestion  flag=vbOkCancel
    If MsgBox(Msg, flag + vbQuestion + vbDefaultButton2) = vbCancel Then
      Err.Raise 110
    End If
  End If
  
End Sub

Private Function minmaj(ByVal val As Double, ByVal signe As Integer) As Double
' Arrondir la valeur avec 4 décimales, selon son signe
  minmaj = Arrond(val, 4) + 0.00049999999 * signe

End Function

Public Sub autorRayon()
Dim i%

  Grille.Col = 2    ' Rayon
  For i = 1 To Grille.MaxRows - 1
    If TabLigneTraj(i).dessinCalculé Then
      Grille.Row = i
      Grille.Lock = True
    End If
  Next
  
End Sub

'************************************************************************************
'
' Propriétés publiques utilisées par frmLargeur
'
'************************************************************************************
Public Property Let SurLarg(ByVal numRemorque As Integer, ByVal surLargeur As Double)
  mSurLarg(numRemorque) = surLargeur
End Property

Public Property Get SurLarg(ByVal numRemorque As Integer) As Double
  SurLarg = mSurLarg(numRemorque)
End Property

Public Property Let Debord(ByVal numRemorque As Integer, ByVal Debordement As Double)
  mDebord(numRemorque) = Debordement
End Property

Public Property Get Debord(ByVal numRemorque As Integer) As Double
  Debord = mDebord(numRemorque)
End Property

'************************************************************************************
' InstrRev n'existe qu'à partir de VB6
'************************************************************************************
Private Function InstrRev(ByVal s1 As String, ByVal s2 As String) As Integer
Dim pos As Integer

Do
  pos = InStr(s1, s2)
  InstrRev = InstrRev + pos
  s1 = Mid(s1, pos + 1)
Loop While pos <> 0

End Function

