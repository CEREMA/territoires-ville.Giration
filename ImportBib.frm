VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImportBib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import de v�hicules"
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
         Caption         =   "Tout d�s�lectionner"
         Height          =   320
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   1600
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Tout s�lectionner"
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
   Begin VB.ListBox lstV�hicules 
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
   Begin VB.Label lblRefus�s 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  v�hicule(s) refus�(s) :  vitesse de braquage sup�rieure � la vitesse maximale admissible"
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

'   R�alisation : Andr� VIGNAUD

'   Module de feuille : frmImportBib   -   IMPORTBIB.FRM
'   Feuille permettant l'import de v�hicules issus d'une biblioth�que de m�me format dans la biblioth�que courante
'**************************************************************************************
Option Explicit
Private BibRenomm�e As New Collection

Const IDm_ExistFichier = "existe d�j�" & vbCrLf & "Voulez-vous le remplacer?"
Const IDm_BibVehIncompatible = "Biblioth�que de v�hicules incompatible"
Const IDm_RenomVehImport = "Renommer le(s) v�hicule(s) avant de l(es) importer"
Const IDm_VehiculeExistant = "Nom de v�hicule existant dans la biblioth�que en cours"
Const IDm_RefusV�hicules = "Aucun v�hicule acceptable"

Private txtFichierAValider As Boolean
Private numEnreg() As Integer
Private numFichVeh As Integer
Private DrapeauBibliVide As Boolean

Private Sub cmdCancel_Click()
  Unload Me
  If DrapeauBibliVide And ColV�hicules.Count = 0 Then Kill GirationVeh
End Sub

'***********************************
' D�selection de toute la liste
'***********************************
Private Sub cmdDeselectAll_Click()
Dim i As Integer

With lstV�hicules
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
' Incorporation dans la biblioth�que des v�hicules s�lectionn�s
'**********************************************************************
Private Sub cmdOK_Click()
Dim numFichLu As Integer
Dim i As Integer
Dim VehicTab As StructV�hicule
Dim NomFich As String
Dim doublon As Boolean
Dim TabRenomm() As Integer
Dim nb As Integer

' Recherche des doublons entre les 2 biblioth�ques (abandon si on en trouve)
  With lstV�hicules
    For i = 0 To .ListCount - 1
      If .Selected(i) And ExistVeh(RTrim(.List(i)), ColV�hicules) Then
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
      ' Biblioth�que courante � enrichir
    Close numFichVeh
    If BibliEnMiseAJour Then Exit Sub
    numFichVeh = FreeFile
    Open GirationVeh For Random Lock Read Write As numFichVeh Len = Len(VehicTab)
      ' Fichier import�
    NomFich = dlgFichier.FileName
    numFichLu = FreeFile
    Open NomFich For Random As numFichLu Len = Len(VehicTab)
      
    On Error GoTo GestErr
    For i = 0 To lstV�hicules.ListCount - 1
      If lstV�hicules.Selected(i) Then
        Get #numFichLu, numEnreg(i), VehicTab
        VehicTab.nom = lstV�hicules.List(i)
        InserV�hicule VehicTab
        Put #numFichVeh, ColV�hicules.Count + 1, VehicTab
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
' Renommer un v�hicule qui a un nom existant d�j� dans la biblioth�que
'***********************************************************************
Private Sub cmdRenommer_Click()
Dim NouveauNom As String

On Error GoTo GestErr
NouveauNom = InputBox("Nouveau nom du v�hicule : " & lstV�hicules.List(lstV�hicules.ListIndex))
If NouveauNom <> "" Then
  If ExistVeh(NouveauNom, ColV�hicules) Then MsgBox IDm_VehiculeExistant: Exit Sub
  BibRenomm�e.Add NouveauNom, NouveauNom
  lstV�hicules.List(lstV�hicules.ListIndex) = NouveauNom
End If

Exit Sub

GestErr:
  If Err = 457 Then ' Un autre v�hicule a d�j� �t� renomm� ainsi
    MsgBox IDm_NomVehicUtilise
  Else
    ErrGeneral
  End If
End Sub

'***********************************
' S�lection de toute la liste
'***********************************
Private Sub cmdSelectAll_Click()
Dim i As Integer

For i = 0 To lstV�hicules.ListCount - 1
  lstV�hicules.Selected(i) = True
Next
lstV�hicules.ListIndex = 0
End Sub

Private Sub Form_Activate()
  If dlgFichier.FileName = "" Then cmdCancel = True
End Sub

'***********************************
' Chargement de la feuille
'***********************************
Private Sub Form_Load()

    'Affichage centr� de la fen�tre
    Me.ScaleMode = 1
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

  ' Verrouilage de la biblioth�que en �criture
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
  Cr�erListeVeh
  lstV�hicules.ListIndex = -1
  Exit Sub
  
GestErr:
  If Err <> cdlCancel Then ErrGeneral
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim nom As Object

For i = 1 To BibRenomm�e.Count
  BibRenomm�e.Remove 1
Next
Close numFichVeh

End Sub

'*********************************************************
' Activation du bouton Renommer si le v�hicule existe d�j�
'**********************************************************
Private Sub lstV�hicules_Click()
Dim i As Integer

With lstV�hicules
  i = .ListIndex
  If i <> -1 Then cmdRenommer.Enabled = .Selected(i) And ExistVeh(RTrim(.List(i)), ColV�hicules)
End With
End Sub

'*****************************************
' L'utilisateur a coch� ou d�coch� un item
'******************************************
Private Sub lstV�hicules_ItemCheck(Item As Integer)
  cmdOK.Enabled = (lstV�hicules.SelCount <> 0)
End Sub

'**************************************************
' Cr�ation de la liste des v�hicules � partir du fichier fourni
'**************************************************
Private Sub Cr�erListeVeh()
Dim numFichLu As Integer
Dim i As Integer
Dim VehicTab As StructV�hicule
Dim NomFich As String
Dim nbRefus�s As Integer

  NomFich = dlgFichier.FileName

  numFichLu = FreeFile
  Open NomFich For Random As numFichLu Len = Len(VehicTab)
  On Error GoTo GestErr
  Get #numFichLu, 1, VehicTab
  If VehicTab.nom <> vehicVersion.nom Then Err.Raise 1001
  For i = 2 To FileLen(NomFich) / Len(VehicTab)
    Get #numFichLu, i, VehicTab
    If Superieur(VehicTab.aVehMax, angConv(OptGen.VitMax, radian)) Then
      nbRefus�s = nbRefus�s + 1
    Else
      ReDim Preserve numEnreg(lstV�hicules.ListCount)
      numEnreg(lstV�hicules.ListCount) = i
      lstV�hicules.AddItem VehicTab.nom
    End If
  Next
  
  If nbRefus�s > 0 Then
    lblRefus�s = CStr(nbRefus�s) & lblRefus�s
    If lstV�hicules.ListCount = 0 Then
      MsgBox lblRefus�s & vbCrLf & IDm_RefusV�hicules, vbOKOnly + vbExclamation
      Unload Me
    Else
      lblRefus�s.Visible = True
    End If
  End If
  Close #numFichLu
  
  Exit Sub
  
GestErr:
    MsgBox IDm_BibVehIncompatible
    Close #numFichLu
    Unload Me
End Sub

