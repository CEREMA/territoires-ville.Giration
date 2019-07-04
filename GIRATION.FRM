VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.MDIForm MDIGiration 
   BackColor       =   &H8000000C&
   Caption         =   "Giration"
   ClientHeight    =   6510
   ClientLeft      =   30
   ClientTop       =   -990
   ClientWidth     =   11400
   Icon            =   "GIRATION.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ilsGiration 
      Left            =   3240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":030A
            Key             =   "curZoom"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":0624
            Key             =   "curPAN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":093E
            Key             =   "curXY"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":0C58
            Key             =   "curDistance"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":0F72
            Key             =   "curDistance2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":128C
            Key             =   "curAngle"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":15A6
            Key             =   "curAngle2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":18C0
            Key             =   "imgGiration"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":1BDA
            Key             =   "imgTraject"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":1EF4
            Key             =   "imgDrapeau"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":220E
            Key             =   "imgVéhicule"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   1320
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgTrajectoire 
      Left            =   720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "gir"
      FilterIndex     =   1
      PrinterDefault  =   0   'False
   End
   Begin MSComctlLib.ImageList ilsToolbar 
      Left            =   2400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2528
            Key             =   "imgNew"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":263A
            Key             =   "imgOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":274C
            Key             =   "imgSave"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":285E
            Key             =   "imgPrint"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2970
            Key             =   "imgZoom"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2A01
            Key             =   "imgZAvant"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2AA7
            Key             =   "imgZArriere"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2B39
            Key             =   "imgPAN"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2C22
            Key             =   "imgFleche"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2F3C
            Key             =   "imgXY"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":2FA5
            Key             =   "imgDistance"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GIRATION.frx":300B
            Key             =   "imgAngle"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrGiration 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ilsToolbar"
      HotImageList    =   "ilsToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnNew"
            Object.ToolTipText     =   "Nouveau"
            ImageKey        =   "imgNew"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnOpen"
            Object.ToolTipText     =   "Ouvrir"
            ImageKey        =   "imgOpen"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnSave"
            Object.ToolTipText     =   "Enregistrer"
            ImageKey        =   "imgSave"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnPrint"
            Object.ToolTipText     =   "Imprimer"
            ImageKey        =   "imgPrint"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnZoom"
            Object.ToolTipText     =   "Zoom"
            ImageKey        =   "imgZoom"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnZAvant"
            Object.ToolTipText     =   "Vue précédente"
            ImageKey        =   "imgZAvant"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnZArriere"
            Object.ToolTipText     =   "Zoom arrière"
            ImageKey        =   "imgZArriere"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnPAN"
            Object.ToolTipText     =   "Panoramique"
            ImageKey        =   "imgPAN"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnCoord"
            Style           =   4
            Object.Width           =   3000
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnFleche"
            Object.ToolTipText     =   "Défilement"
            ImageKey        =   "imgFleche"
            Value           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnXY"
            Object.ToolTipText     =   "Coordonnées"
            ImageKey        =   "imgXY"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnDistance"
            Object.ToolTipText     =   "Distance"
            ImageKey        =   "imgDistance"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btnAngle"
            Object.ToolTipText     =   "Angle"
            ImageKey        =   "imgAngle"
         EndProperty
      EndProperty
      Begin MSComctlLib.StatusBar staBarre 
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Text            =   "x="
               TextSave        =   "x="
               Key             =   "X"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Text            =   "y="
               TextSave        =   "y="
               Key             =   "Y"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar staMesure 
         Height          =   375
         Left            =   7920
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
               Key             =   "Distance"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuZoomMul 
      Caption         =   "&Zoom"
      Visible         =   0   'False
      Begin VB.Menu mnuZoom 
         Caption         =   "Fenêtre"
         Index           =   0
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Panoramique"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Vue précédente"
         Index           =   2
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom arrière"
         Index           =   3
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Tout voir"
         Index           =   4
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Sortir"
         Index           =   6
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "&Fichier"
      Index           =   0
      Begin VB.Menu mnuFich 
         Caption         =   "&Nouvelle trajectoire"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFich 
         Caption         =   "&Ouvrir une trajectoire ..."
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFich 
         Caption         =   "&Fermer"
         Index           =   3
      End
      Begin VB.Menu mnuFich 
         Caption         =   "&Enregistrer la trajectoire"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Enregistrer la trajectoire &sous ..."
         Index           =   5
      End
      Begin VB.Menu mnuFich 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFich 
         Caption         =   "E&xporter l'épure ..."
         Index           =   7
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Exporter &les caractéristiques..."
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu mnuFich 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Bibliothèque de &véhicules"
         Index           =   10
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Exporter la &bibliothèque..."
         Index           =   11
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Importer &une bibliothèque..."
         Enabled         =   0   'False
         Index           =   12
      End
      Begin VB.Menu mnuFich 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFich 
         Caption         =   "&Importer un fond de plan ..."
         Index           =   14
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Enregistrer le fon&d de plan sous ..."
         Index           =   15
      End
      Begin VB.Menu mnuFich 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuFich 
         Caption         =   "&Configuration de l'impression ..."
         Index           =   17
      End
      Begin VB.Menu mnuFich 
         Caption         =   "Im&primer ..."
         Index           =   18
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFich 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "&Outils"
      Index           =   1
      Begin VB.Menu mnuOutils 
         Caption         =   "&Insérer un déplacement"
         Index           =   1
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Supprimer un déplacement"
         Index           =   2
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Décomposer un Virage"
         Index           =   3
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Pas à pas"
         Index           =   5
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "Sur&largeur"
         Index           =   6
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "-"
         HelpContextID   =   36
         Index           =   7
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Cadrage"
         Index           =   8
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Rafraîchir"
         Index           =   9
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Options générales"
         Index           =   11
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "Options d'&affichage"
         Index           =   12
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Mesures"
         Index           =   14
      End
      Begin VB.Menu mnuOutils 
         Caption         =   "&Totalisation des angles"
         Index           =   15
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "Fe&nêtre"
      Index           =   2
      WindowList      =   -1  'True
      Begin VB.Menu mnuFenetre 
         Caption         =   "&Cascade"
         Index           =   1
      End
      Begin VB.Menu mnuFenetre 
         Caption         =   "Mosaïque &horizontale"
         Index           =   2
      End
      Begin VB.Menu mnuFenetre 
         Caption         =   "&Réorganiser les icônes"
         Index           =   3
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "&?"
      Index           =   3
      Begin VB.Menu mnuAide 
         Caption         =   "&Sommaire"
         Index           =   0
      End
      Begin VB.Menu mnuAide 
         Caption         =   "A&ide sur..."
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAide 
         Caption         =   "&Rechercher..."
         Index           =   2
      End
      Begin VB.Menu mnuAide 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAide 
         Caption         =   "A &propos de Giration ..."
         Index           =   4
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "&Licence"
      End
   End
End
Attribute VB_Name = "MDIGiration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille MDI : MDIGiration   -   GIRATION.FRM
'   Feuille Maitresse (MDI) de GIRATION

'**************************************************************************************
Option Explicit
                         
Private Sub MDIForm_Activate()
      'Modif v3.3(25/04/07) : Utilisation d'un énumérateur au lieu de constantes
'  If Not fCourante Is Nothing And gtOutil >= 0 And gtOutil < 4 Then
  If Not fCourante Is Nothing And gtOutil <> TOOL_AUCUN Then
    tbrGiration.Buttons(gtBoutons(gtOutil)).Value = tbrPressed
  End If
End Sub

  Private Sub MDIForm_Load()
    
    On Error GoTo GestErr
    
    If FichierJournal Then Write #numFichLog, "Ouverture de GIRATION"
    
    tbrGiration.Buttons("btnCoord").Width = staBarre.Width
    If VersionDemo Then Caption = App.Title & " - " & ID_ExemplDemo
    
    HelpContextID = SOMMAIRE '0
        
    'Affichage centré de la fenêtre
    Me.Height = Screen.Height - 100 ' 100 : Réservation pour la barre des tâches
    Me.Width = Screen.Width
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    'Nombre d'instance de la fenêtre frmTraject égal à zéro par défaut.
    NbreFilles = 0
   
    'Affichage des menus
    GriserMenus Etat:=False
      
    'Paramétrage aide contextuelle
    App.HelpFile = App.Path + "\" + HELPNAME
    dlgTrajectoire.HelpFile = App.HelpFile
    dlgTrajectoire.Filter = Majus(ID_Trajectoire) & " (*.gir)|*.gir|" & Majus(ID_Trajectoire) & " Version2 (*.egt)|*.egt"
    
    'Paramétrage de l'aide
    mnuFich.Item(1).HelpContextID = NOUVTRAJ
    mnuFich.Item(2).HelpContextID = OUVRIRTRAJ
    mnuFich.Item(4).HelpContextID = ENREGTRAJ
    mnuFich.Item(5).HelpContextID = ENREGTRAJ
    mnuFich.Item(7).HelpContextID = EXPORTEPURE
    mnuFich.Item(8).HelpContextID = EXPORTCARAC
    mnuFich.Item(10).HelpContextID = BIBVEHICUL
    mnuFich.Item(11).HelpContextID = EXPORTVEHICUL
    mnuFich.Item(12).HelpContextID = IMPORTVEHICUL
    mnuFich.Item(14).HelpContextID = IMPORTFDP
    mnuFich.Item(15).HelpContextID = ENREGFDP
    mnuFich.Item(17).HelpContextID = 0
    mnuFich.Item(18).HelpContextID = IMPRIMER

    mnuOutils.Item(1).HelpContextID = INSERSUPPRDEPL
    mnuOutils.Item(2).HelpContextID = INSERSUPPRDEPL
    mnuOutils.Item(3).HelpContextID = DECOMPVIRAGE
    mnuOutils.Item(5).HelpContextID = PASAPAS
    mnuOutils.Item(6).HelpContextID = LARGEUR
    mnuOutils.Item(8).HelpContextID = CADRAGE
    mnuOutils.Item(11).HelpContextID = OPTGENERAL
    mnuOutils.Item(12).HelpContextID = OPTAFFICHAGE
    mnuOutils.Item(14).HelpContextID = MESURES
    mnuOutils.Item(15).HelpContextID = ANGLETOTAL
    
    'Mise à jour de l'ihm du à QLM
    Call InitQlm
    Exit Sub
    
GestErr:
  ErrGeneral
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim i%
Dim MySettings As Variant

  If FichierJournal Then
    Write #numFichLog, "Fin de GIRATION"
    Close #numFichLog
  End If

  On Error GoTo GestErr

  ReInitializePrinter

  SaveSetting Appname:=App.Title, SECTION:="Version", _
    Key:="V", Setting:=GirationVersion
  SaveSetting Appname:=App.Title, SECTION:="Parametres", _
    Key:="Mesure", Setting:=mnuOutils(MNUMESURES).Checked
  SaveSetting Appname:=App.Title, SECTION:="Parametres", _
    Key:="TotalAngle", Setting:=mnuOutils(MNUANGLETOTAL).Checked
  
  For i = 0 To UBound(MRUFichiers)
    SaveSetting Appname:=App.Title, SECTION:="Recent Files", _
    Key:="File" & CStr(i + 1), Setting:=MRUFichiers(i)
  Next
  ' Suppression dans la registry des fichiers effacés (à reprendre en même temps que MRUmenu)
  MySettings = GetAllSettings(Appname:=App.Title, SECTION:="Recent Files")
  If Not IsEmpty(MySettings) Then
    For i = UBound(MRUFichiers) + 1 To UBound(MySettings, 1)
      DeleteSetting App.Title, "Recent Files", MySettings(i, 0)
    Next
  End If

'Indispensable pour que la procédure Main s'arrête si erreur de protection
  End

GestErr:
  'Err=9 - Ubound est en erreur, car on n'a pas encore initialisé le projet
  If Err <> 9 Then
    MsgBox IDm_ErrFatale & CStr(Err) & vbCrLf & Err.Description, vbOKOnly + vbCritical
  End If
  
  End
            
End Sub

Private Sub mnuFich_Click(Index As Integer)
' en retour, flag reçoit True si l'enregistrement est abandonné - Ne sert pas dans ce menu
Dim flag As Integer

If EstChargée(frmPas) Then frmPas.SetFocus: Exit Sub

          'Menu fichiers
          
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
    'lireProtect
          
    Select Case Index
        Case MNUNOUVTRAJ 'Nouvelle trajectoire
          lireVeh Relecture:=True
          If ColVéhicules.Count = 0 Then
            MsgBox (IDm_BibVehVide)
          Else
              frmTrajpar.Show vbModal
          End If
        
        Case MNUOUVRIRTRAJ 'Ouvrir une trajectoire
          Ouvrir
        
        Case 3 'Fermer (la trajectoire courante)
          Unload fCourante
        
        Case MNUENREGTRAJ 'Enregistrer la trajectoire
          Enregistrer fCourante, flag             ' en retour, flag reçoit True si l'enregistrement est abandonné
            
        Case MNUENREGTRAJSOUS 'Enregistrer la trajectoire sous
          EnregSous fCourante, flag             ' en retour, flag reçoit True si l'enregistrement est abandonné
        
        Case MNUEXPORTEPURE  'Exporter une trajectoire
          Export
          
        Case MNUEXPORTCARAC
          fCourante.ecrireFormatTexte
          
        Case MNUBIBVEH 'Bibliothèque de véhicules
          lireVeh Relecture:=True
          frmBibVéhicule.Show vbModal
                    
        Case MNUIMPORTBIB
          If Not BibliEnMiseAJour Then
            frmImportBib.Show vbModal
          Else
            gtBibliVerrouillée = False
          End If
            
        Case MNUEXPORTBIB
          lireVeh Relecture:=True
          If ColVéhicules.Count = 0 Then
            MsgBox (IDm_BibVehVide)
          Else
            BibliEnMiseAJour premier:=True
            gtBibliVerrouillée = False
            frmExportBib.Show vbModal
          End If
            
        Case MNUIMPORTFDP 'Importer un fond de plan
          Dim Cancel As Integer

          If fCourante.FDP <> "" Then
            If MsgBox(IDm_RemplacerFDP & " " & fCourante.FDP & " ?", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
            If fCourante.dlgFDP.FileName = "" Then
              Select Case MsgBox(IDm_EnregistrerFDP & " " & fCourante.Caption & " ?", vbYesNoCancel + vbQuestion)
              Case vbYes
                EnregFDPSous fCourante, Cancel
              Case vbCancel
                Cancel = True
              End Select
            End If
          End If
          
          If Not Cancel Then
            fCourante.dlgFDP.FileName = ""
            Import
          End If
            
        
        Case MNUENREGFDP 'Enregistrer le fond de plan sous
              
          EnregFDPSous fCourante, flag
            
        Case MNUCONFIGIMPR 'Configuration de l'impression
      
        ShowPrinter Me
        
''            On Error GoTo ErrHandler
''
''            dlgPrint.flags = cdlPDPrintSetup   ' Pour afficher directement la fenêtre Configuration
''            dlgPrint.ShowPrinter
    
        Case MNUIMPRIMER 'Imprimer
            
            frmImprim.Show vbModal
            
    End Select
    
    If Screen.ActiveForm Is Nothing And Not fCourante Is Nothing Then fCourante.Show
    
       
''    Exit Sub
''
''ErrHandler:
''    If Err <> cdlCancel Then    ' 32755
''      MsgBox IDm_ErrImprim & " (" & Format(Err.Number, "000") & vbCrLf & Err.Description
''    End If
''   Exit Sub

End Sub

Private Sub mnuLicence_Click()
     frmKey.Show vbModal
    'Mise à jour de l'ihm
     InitQlm
End Sub

Private Sub mnuQuit_Click()
            'Quitter
            'Quitte l'application
            Unload Me

End Sub

Private Sub mnuOutils_Click(Index As Integer)
  'Menu Outils
    
    If Not gtInteractionEnCours Then
  
'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
    'lireProtect
    
    Select Case Index
        Case MNUINSERDEPL 'Insérer un déplacement
          fCourante.InserDepl
           
        Case MNUSUPPRDEPL 'Supprimer un déplacement
            fCourante.SupprDepl
            
        Case MNUDECOMPVIRAGE 'Décomposer un virage
            'Possible si une ligne virage est sélectionnée dans le spread ,,,
              fCourante.DECOMPVIRAGE
            
        Case MNUPASAPAS 'Pas à pas
            If fCourante.flagRecalcul Then Rafraichir
            frmPas.Show vbModeless, Me
            
        Case MNULARGEUR 'Surlargeur
            frmLargeur.Show vbModal
        
        Case MNUCADRAGE 'Cadrage
'            frmCadrage.ContexteHorsBouton = True
            frmCadrage.Show vbModal
        
        Case MNURAFRAICHIR 'Rafraichir
          Rafraichir
        
        Case MNUOPTGENERAL 'Options générales
          frmOptGen.Show vbModal
    
        Case MNUOPTAFFICHAGE 'Options d'affichage
          frmOptAff.Show vbModal
          
        Case MNUMESURES ' Mesures
          gtCoordActif = mnuOutils(Index).Checked Or gtOutil = TOOL_COORDDYN
          mnuOutils(Index).Checked = Not mnuOutils(Index).Checked
          DefautCurseur
          If Not fCourante Is Nothing Then fCourante.Form_Activate
    
        Case MNUANGLETOTAL ' Angle total
          mnuOutils(Index).Checked = Not mnuOutils(Index).Checked
          If Not fCourante Is Nothing Then fCourante.fraAngleTotal.Visible = mnuOutils(Index).Checked And fCourante.Grille.Visible
    End Select

    If Screen.ActiveForm Is Nothing And Not fCourante Is Nothing Then fCourante.Show
  End If

End Sub

Private Sub mnuFenetre_Click(Index As Integer)
    'Menu fenêtre

If gtInteractionEnCours Then Exit Sub
If EstChargée(frmPas) Then frmPas.SetFocus: Exit Sub

Select Case Index
        Case 1 'Cascade
'            MDIGiration.Arrange vbCascade
            Arrange vbCascade
            
        Case 2  ' Mosaïque horizontale
            MDIGiration.Arrange vbTileHorizontal
        
        Case 3 'Pour les icônes
            MDIGiration.Arrange vbArrangeIcons
    End Select
End Sub

Private Sub mnuSelect_Click(Index As Integer)
' Ouverture d'un fichier .GIR par choix dans la liste des derniers utilisés
Dim NomFich As String

If gtInteractionEnCours Then Exit Sub
If EstChargée(frmPas) Then frmPas.SetFocus: Exit Sub

'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
'lireProtect

NomFich = MRUFichiers(Index)
If ExistFich(NomFich) Then
  dlgTrajectoire.FileName = NomFich
  creFille nouveau:=False
Else      ' En principe, çà ne devrait pas arriver, le controle d'existence étant fait au chargement de GIRATION dans MRUMenu (GIRATION.BAS)
  MsgBox NomFich & vbCrLf & IDm_MRUFichierDisparu, vbExclamation
End If

End Sub

Private Sub mnuAide_Click(Index As Integer)
    Dim Retour%
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    
    Select Case Index
        Case 0 'Aide Sommaire
          'Modif AV : 08/11/2006
          'L'onglet Index étant supprimé de l'aide en ligne, il vaut mieux suprimer l'item de menu Aide sur...
          'Du coup
          ' - on affecte la touche F1 à l'item 0 au lieu de l'item 1
          ' - on rend invisible l'item 1
          ' - et on gère contextuellement l'index 0
          Call objHelp.Show(App.HelpFile, "Main")
          
'          If HelpContextID = 0 Then
'            'Sommaire
'            Call objHelp.Show(App.HelpFile, "Main")
'          Else
'            'Aide contextuelle
'            Call objHelp.Show(App.HelpFile, "Main", HelpContextID)
'          End If

        Case 1 'Aide sur
            'Modif fait par Frank Trifiletti on utilise le contextid de la fenêtre étude en cours
            'qui est dans la globale monetude dont son helpcontextid est mis à jour dans la sub ChangerHelpId
            'qui est appellé à chaque Form_Activate et dans le TabData_Click de frmDocument.frm
            'car le contextid était toujours nulle avec showindex normal on ne le passe pas en argument.
            If fCourante Is Nothing Then
                'Cas d'appel  de F1 si aucun étude ouverte sinon plantage
                'Onglet Index supprimé!!!
                'Call objHelp.ShowIndex(App.HelpFile, "Main")
                Call objHelp.Show(App.HelpFile, "Main")
            Else
                Call objHelp.Show(App.HelpFile, "Main", fCourante.HelpContextID)
            End If
            'Fin modif F.Trifiletti
                
        Case 2 'Aide Rechercher
            Call objHelp.ShowSearch(App.HelpFile, "Main")
          
        Case 4 'A propos de Giration
           frmApropos.Show vbModal
    
    End Select
    Set objHelp = Nothing
End Sub


Public Sub mnuZoom_Click(Index As Integer)
Dim i As Integer

' On désarme les flags pour que gtInteractionEnCours ne refuse pas
  fCourante.flagZoom = False
  fCourante.flagPAN = False
  For i = TOOL_ZOOM To TOOL_ZTOUT
    mnuZoom(i).Checked = (i = Index)
  Next
  
  Select Case Index
  Case TOOL_ZOOM
    fCourante.cmdZoom = True
  Case TOOL_PAN
    fCourante.cmdPAN = True
  Case TOOL_ZAVANT
    fCourante.cmdZAvant = True
    DoEvents
    If fCourante.Milieux.Count = 1 Then SendKeys "{ESC}"
  Case TOOL_ZARRIERE
    fCourante.cmdZArriere = True
  Case TOOL_ZTOUT
    'Modif v3.3(02/05/07) : On évite ainsi de charger frmCadrage inutilement
    'frmCadrage.cmdEtendu = True
    GirStand.ExecuteZoomEtendu
  Case Else
    SendKeys "{ESC}"
  End Select
    
End Sub

Private Sub tbrGiration_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As TOOLEnum
Dim ActiverBouton As Boolean

ActiverBouton = True

Select Case Button.Key
Case "btnNew"
  mnuFich_Click MNUNOUVTRAJ
Case "btnOpen"
  mnuFich_Click MNUOUVRIRTRAJ
Case "btnSave"
  mnuFich_Click MNUENREGTRAJ
Case "btnPrint"
  mnuFich_Click MNUIMPRIMER
Case "btnZoom"
  fCourante.cmdZoom = True
Case "btnZAvant"
  fCourante.cmdZAvant = True
Case "btnPAN"
  fCourante.cmdPAN = True
Case "btnZArriere"
  fCourante.cmdZArriere = True
Case Else
  If gtInteractionEnCours Then Exit Sub
  Select Case gtOutil
  Case TOOL_COORDDYN
    If Button.Key = "btnFleche" Then gtOutil = TOOL_AUCUN: ActiverBouton = False
  Case TOOL_XY
    If Button.Key = "btnXY" Then gtOutil = TOOL_AUCUN: ActiverBouton = False
  Case TOOL_DIST
    If Button.Key = "btnDistance" Then gtOutil = TOOL_AUCUN: ActiverBouton = False
  Case TOOL_ANGLEDYN
    If Button.Key = "btnAngle" Then gtOutil = TOOL_AUCUN: ActiverBouton = False
  End Select

  If ActiverBouton Then
    For i = TOOL_AUCUN + 1 To TOOL_ANGLEDYN
      If Button.Key = gtBoutons(i) Then
        Button.Value = tbrPressed
        gtOutil = i
      Else
        tbrGiration.Buttons(gtBoutons(i)).Visible = False
        tbrGiration.Buttons(gtBoutons(i)).Value = tbrUnpressed
        tbrGiration.Buttons(gtBoutons(i)).Visible = True
      End If
      DoEvents
    Next
  End If
  
  fCourante.flagPAN = False
  InitialOutilMesure
End Select

End Sub

Private Sub InitialOutilMesure()
    'Mise en commentaire v3.3(25/04/07) : DéfautCurseur ne charge plus frmCadrage (implicitement par cmdDésigner)
    'If EstChargée(frmCadrage) Then Unload frmCadrage
    If gtOutil = TOOL_COORDDYN Then
      gtCoordActif = True
    Else
      gtCoordActif = False
      staBarre.Panels("X").Text = ""
      staBarre.Panels("Y").Text = ""
    End If
    staMesure.Panels("Distance").Text = ""
    staMesure.Visible = (gtOutil = TOOL_DIST Or gtOutil = TOOL_ANGLEDYN)
    DefautCurseur
End Sub

Private Sub InitQlm()
  'Initialisation des menus modifiés par QLM
  'les variables globales sont maj par protection.bas
  'ATTENTION : vérifier les noms des menus!!!
  Me.mnuHelpBar2.Visible = GvisibiliteMnuBarre
  Me.mnuLicence.Visible = GvisibiliteMnuLicence
  'a adapter en fonction du clogiciel
'  Me.Caption = "Giration v" + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision) + GmodifTitreApplication
  Me.Caption = "Giration " + Format(App.Major) + "." + Format(App.Minor) + GmodifTitreApplication
  'fin initialisation qlm
  'fin initialisation qlm
End Sub
