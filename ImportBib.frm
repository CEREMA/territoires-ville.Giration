VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImportBib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import de véhicules"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "ImportBib.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBoutons 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   3690
      ScaleHeight     =   4665
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuler"
         CausesValidation=   0   'False
         Height          =   320
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1600
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   320
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1600
      End
      Begin VB.CommandButton cmdDeselectAll 
         Caption         =   "Tout désélectionner"
         Height          =   320
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   1600
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Tout sélectionner"
         Height          =   320
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   1600
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Aide"
         Height          =   320
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   1600
      End
      Begin VB.CommandButton cmdRenommer 
         Caption         =   "Renommer"
         Enabled         =   0   'False
         Height          =   320
         Left            =   120
         TabIndex        =   0
         Top             =   1200
         Width           =   1600
      End
   End
   Begin VB.ListBox lstVéhicules 
      Height          =   3660
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog dlgFichier 
      Left            =   6960
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblRefusés 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  véhicule(s) refusé(s) :  vitesse de braquage supérieure à la vitesse maximale admissible"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmImportBib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3.2 - CERTU/CETE de l'Ouest
'         Juin 2000

'   Réalisation : André VIGNAUD

'   Module de feuille : frmImportBib   -   IMPORTBIB.FRM
'   Feuille permettant l'import de véhicules issus d'une bibliothèque de même format dans la bibliothèque courante
'**************************************************************************************
Option Explicit
Private BibRenommée As New Collection

Const IDm_ExistFichier = "existe déjà" & vbCrLf & "Voulez-vous le remplacer?"
Const IDm_BibVehIncompatible = "Bibliothèque de véhicules incompatible"
Const IDm_RenomVehImport = "Renommer le(s) véhicule(s) avant de l(es) importer"
Const IDm_VehiculeExistant = "Nom de véhicule existant dans la bibliothèque en cours"
Const IDm_RefusVéhicules = "Aucun véhicule acceptable"

Private txtFichierAValider As Boolean
Private numEnreg() As Integer
Private numFichVeh As Integer
Private DrapeauBibliVide As Boolean

Private Sub cmdCancel_Click()
  Unload Me
  If DrapeauBibliVide And ColVéhicules.Count = 0 Then Kill GirationVeh
End Sub

'***********************************
' Déselection de toute la liste
'***********************************
Private Sub cmdDeselectAll_Click()
Dim i As Integer

With lstVéhicules
  For i = 0 To .ListCount - 1
    .Selected(i) = False
  Next
  .ListIndex = -1
End With

End Sub

Private Sub cmdHelp_Click()
  SendKeys "{F1}", True
End Sub

'**********************************************************************
' Incorporation dans la bibliothèque des véhicules sélectionnés
'**********************************************************************
Private Sub cmdOK_Click()
Dim numFichLu As Integer
Dim i As Integer
Dim VehicTab As StructVéhicule
Dim NomFich As String
Dim doublon As Boolean
Dim TabRenomm() As Integer
Dim nb As Integer

' Recherche des doublons entre les 2 bibliothèques (abandon si on en trouve)
  With lstVéhicules
    For i = 0 To .ListCount - 1
      If .Selected(i) And ExistVeh(RTrim(.List(i)), ColVéhicules) Then
        .Selected(i) = True
        If Not doublon Then .TopIndex = i: .ListIndex = i
        doublon = True
      End If
    Next
  End With
  If doublon Then
    MsgBox IDm_RenomVehImport
    Exit Sub
  End If
  
  'Incorporation effective
      ' Bibliothèque courante à enrichir
    Close numFichVeh
    If BibliEnMiseAJour Then Exit Sub
    numFichVeh = FreeFile
    Open GirationVeh For Random Lock Read Write As numFichVeh Len = Len(VehicTab)
      ' Fichier importé
    NomFich = dlgFichier.FileName
    numFichLu = FreeFile
    Open NomFich For Random As numFichLu Len = Len(VehicTab)
      
    On Error GoTo GestErr
    For i = 0 To lstVéhicules.ListCount - 1
      If lstVéhicules.Selected(i) Then
        Get #numFichLu, numEnreg(i), VehicTab
        VehicTab.nom = lstVéhicules.List(i)
        InserVéhicule VehicTab
        Put #numFichVeh, ColVéhicules.Count + 1, VehicTab
      End If
    Next
    
    Close #numFichLu
    
    Unload Me
    
    Exit Sub
  
GestErr:
    If Err = 457 Then
      doublon = True
      Resume Next
    Else
      MsgBox IDm_BibVehIncompatible
    End If
End Sub


'***********************************************************************
' Renommer un véhicule qui a un nom existant déjà dans la bibliothèque
'***********************************************************************
Private Sub cmdRenommer_Click()
Dim NouveauNom As String

On Error GoTo GestErr
NouveauNom = InputBox("Nouveau nom du véhicule : " & lstVéhicules.List(lstVéhicules.ListIndex))
If NouveauNom <> "" Then
  If ExistVeh(NouveauNom, ColVéhicules) Then MsgBox IDm_VehiculeExistant: Exit Sub
  BibRenommée.Add NouveauNom, NouveauNom
  lstVéhicules.List(lstVéhicules.ListIndex) = NouveauNom
End If

Exit Sub

GestErr:
  If Err = 457 Then ' Un autre véhicule a déjà été renommé ainsi
    MsgBox IDm_NomVehicUtilise
  Else
    ErrGeneral
  End If
End Sub

'***********************************
' Sélection de toute la liste
'***********************************
Private Sub cmdSelectAll_Click()
Dim i As Integer

For i = 0 To lstVéhicules.ListCount - 1
  lstVéhicules.Selected(i) = True
Next
lstVéhicules.ListIndex = 0
End Sub

Private Sub Form_Activate()
  If dlgFichier.FileName = "" Then cmdCancel = True
End Sub

'***********************************
' Chargement de la feuille
'***********************************
Private Sub Form_Load()

    'Affichage centré de la fenêtre
    Me.ScaleMode = 1
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

  ' Verrouilage de la bibliothèque en écriture
    DrapeauBibliVide = Not ExistFich(GirationVeh)
    numFichVeh = FreeFile
    Open GirationVeh For Random Lock Write As numFichVeh

  ' Aide contextuelle
    Me.HelpContextID = IMPORTVEHICUL
    

On Error GoTo GestErr
  With dlgFichier
    .Filter = ID_Bibveh & " (*.veh)|*.veh|"
    .InitDir = App.Path
    .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    .ShowOpen
  End With
  CréerListeVeh
  lstVéhicules.ListIndex = -1
  Exit Sub
  
GestErr:
  If Err <> cdlCancel Then ErrGeneral
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim nom As Object

For i = 1 To BibRenommée.Count
  BibRenommée.Remove 1
Next
Close numFichVeh

End Sub

'*********************************************************
' Activation du bouton Renommer si le véhicule existe déjà
'**********************************************************
Private Sub lstVéhicules_Click()
Dim i As Integer

With lstVéhicules
  i = .ListIndex
  If i <> -1 Then cmdRenommer.Enabled = .Selected(i) And ExistVeh(RTrim(.List(i)), ColVéhicules)
End With
End Sub

'*****************************************
' L'utilisateur a coché ou décoché un item
'******************************************
Private Sub lstVéhicules_ItemCheck(Item As Integer)
  cmdOK.Enabled = (lstVéhicules.SelCount <> 0)
End Sub

'**************************************************
' Création de la liste des véhicules à partir du fichier fourni
'**************************************************
Private Sub CréerListeVeh()
Dim numFichLu As Integer
Dim i As Integer
Dim VehicTab As StructVéhicule
Dim NomFich As String
Dim nbRefusés As Integer

  NomFich = dlgFichier.FileName

  numFichLu = FreeFile
  Open NomFich For Random As numFichLu Len = Len(VehicTab)
  On Error GoTo GestErr
  Get #numFichLu, 1, VehicTab
  If VehicTab.nom <> vehicVersion.nom Then Err.Raise 1001
  For i = 2 To FileLen(NomFich) / Len(VehicTab)
    Get #numFichLu, i, VehicTab
    If Superieur(VehicTab.aVehMax, angConv(OptGen.VitMax, radian)) Then
      nbRefusés = nbRefusés + 1
    Else
      ReDim Preserve numEnreg(lstVéhicules.ListCount)
      numEnreg(lstVéhicules.ListCount) = i
      lstVéhicules.AddItem VehicTab.nom
    End If
  Next
  
  If nbRefusés > 0 Then
    lblRefusés = CStr(nbRefusés) & lblRefusés
    If lstVéhicules.ListCount = 0 Then
      MsgBox lblRefusés & vbCrLf & IDm_RefusVéhicules, vbOKOnly + vbExclamation
      Unload Me
    Else
      lblRefusés.Visible = True
    End If
  End If
  Close #numFichLu
  
  Exit Sub
  
GestErr:
    MsgBox IDm_BibVehIncompatible
    Close #numFichLu
    Unload Me
End Sub

