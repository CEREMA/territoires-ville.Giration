VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportBib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export de véhicules"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "ExportBib.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBoutons 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   3675
      ScaleHeight     =   3705
      ScaleWidth      =   1815
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmdDeselectAll 
         Caption         =   "Tout désélectionner"
         Height          =   320
         Left            =   0
         TabIndex        =   1
         Top             =   1680
         Width           =   1600
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Tout sélectionner"
         Height          =   320
         Left            =   0
         TabIndex        =   0
         Top             =   1200
         Width           =   1600
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   320
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   1600
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuler"
         CausesValidation=   0   'False
         Height          =   320
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   1600
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Aide"
         Height          =   320
         Left            =   0
         TabIndex        =   4
         Top             =   2160
         Width           =   1600
      End
   End
   Begin VB.ListBox lstVéhicules 
      Height          =   3435
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   240
      Width           =   3255
   End
   Begin VB.PictureBox picFichier 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5490
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3705
      Visible         =   0   'False
      Width           =   5490
      Begin VB.CommandButton cmdParcourir 
         Caption         =   "&Parcourir..."
         Height          =   420
         Left            =   5880
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtFichier 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog dlgFichier 
      Left            =   6960
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "Export.veh"
   End
End
Attribute VB_Name = "frmExportBib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************************************
'     GIRATION v3.2 - CERTU/CETE de l'Ouest
'         Juin 2000

'   Réalisation : André VIGNAUD

'   Module de feuille : frmExportBib   -   EXPORTBIB.FRM
'   Feuille permettant d'exorter des véhicules depuis la bibliothèque courante
'**************************************************************************************
Option Explicit

Const IDm_SaisirFichier = "Saisir un nom de fichier"
Const IDm_ExistFichier = "existe déjà" & vbCrLf & "Voulez-vous le remplacer?"
Const IDm_BibvehNonPerso = "La bibliothèque ne contient aucun véhicule personnalisé"
Const IDm_NomBibvehReservé = GIRATIONVEHCOURT & " est un nom de bibliothèque réservé par GIRATION"

Private txtFichierAValider As Boolean

Private Sub cmdCancel_Click()
  Unload Me
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
' Création de la bibliothèque à partir des véhicules sélectionnés
'**********************************************************************
Private Sub cmdOK_Click()
Dim numFich As Integer
Dim i As Integer
Dim VehicTab As StructVéhicule
Dim Abandon As Boolean
Dim numEnreg As Integer

    cmdParcourir = True

    If txtFichier <> "" Then
      numFich = FreeFile
      If ExistFich(txtFichier) Then Kill txtFichier
      Open txtFichier For Random As numFich Len = Len(VehicTab)
      numEnreg = 1
      Put #numFich, numEnreg, vehicVersion
      For i = 0 To lstVéhicules.ListCount - 1
        If lstVéhicules.Selected(i) Then
            'Mise au format StructVéhicule l'objet Véhicule
          ConvObjetStruct ColVéhicules(UCase(lstVéhicules.List(i))), VehicTab
            'Ecriture du véhicule dans le fichier
          numEnreg = numEnreg + 1
          Put #numFich, numEnreg, VehicTab
        End If
      Next
      
      Close #numFich
      Unload Me
   End If
    
    
End Sub

'******************************************************************************
' Navigateur pour sortie sur fichier
'******************************************************************************
Private Sub cmdParcourir_Click()
Dim i As Integer

  With dlgFichier
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    If .InitDir = "" Then
      .InitDir = App.Path
    End If
    .Filter = ID_Bibveh & " (*.veh)|*.veh|"
    .FileName = "Export.veh"
    
    On Error GoTo GestErr
     .ShowSave
     
    If .FileName <> "" Then
      If FichierProtégé(.FileName) Then Exit Sub
      txtFichier = .FileName
      If StrComp(nomCourt(txtFichier, SansExtension:=False), GIRATIONVEHCOURT, vbTextCompare) = 0 Then
        MsgBox IDm_NomBibvehReservé
        txtFichier = ""
        Exit Sub
      End If
      txtFichierAValider = False
    End If
  End With
  
  Exit Sub
  
GestErr:
  If Err = cdlCancel Then Exit Sub ' Else ErreurFatale
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
  If lstVéhicules.ListCount = 0 Then
    Hide
    MsgBox IDm_BibvehNonPerso
    Unload Me
  End If
End Sub

'***********************************
' Chargement de la feuille
'***********************************
Private Sub Form_Load()
Dim Vehic As VEHICULE

    'Affichage centré de la fenêtre
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
  
  ' Aide contextuelle
    Me.HelpContextID = EXPORTVEHICUL
    
  For Each Vehic In ColVéhicules
    'Création de la table de correspondance entre le numéro de véhicule dans la collection et l'item dans la listbox
    If Not Vehic.Protégé Then
    ' Seuls les véhicules personnalisés sont exportables
      lstVéhicules.AddItem Vehic.Nom
    End If
  Next
  lstVéhicules.ListIndex = -1
End Sub

'*****************************************
' L'utilisateur a coché ou décoché un item
'******************************************
Private Sub lstVéhicules_ItemCheck(Item As Integer)
  cmdOK.Enabled = (lstVéhicules.SelCount <> 0)
End Sub


Private Sub txtFichier_Change()
  txtFichierAValider = True
End Sub

'*****************************************
' Validation du nom de fichier à créer
'******************************************
Private Sub txtFichier_Validate(Cancel As Boolean)
  If txtFichier = "" Then Exit Sub
  
  If Extension(txtFichier) = "" Then txtFichier = txtFichier + dlgFichier.DefaultExt
  
  If StrComp(nomCourt(txtFichier, SansExtension:=False), GIRATIONVEHCOURT, vbTextCompare) = 0 Then
    MsgBox IDm_NomBibvehReservé
    Cancel = True
    Exit Sub
  End If
  
  If ExistFich(txtFichier) Then
    If MsgBox(txtFichier & " " & IDm_ExistFichier, vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
      txtFichier.SetFocus
      txtFichier.SelStart = 0
      txtFichier.SelLength = Len(txtFichier)
      Cancel = True
    Else
      txtFichierAValider = False
    End If
  End If
End Sub

