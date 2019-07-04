VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportBib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export de v�hicules"
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
         Caption         =   "Tout d�s�lectionner"
         Height          =   320
         Left            =   0
         TabIndex        =   1
         Top             =   1680
         Width           =   1600
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Tout s�lectionner"
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
   Begin VB.ListBox lstV�hicules 
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

'   R�alisation : Andr� VIGNAUD

'   Module de feuille : frmExportBib   -   EXPORTBIB.FRM
'   Feuille permettant d'exorter des v�hicules depuis la biblioth�que courante
'**************************************************************************************
Option Explicit

Const IDm_SaisirFichier = "Saisir un nom de fichier"
Const IDm_ExistFichier = "existe d�j�" & vbCrLf & "Voulez-vous le remplacer?"
Const IDm_BibvehNonPerso = "La biblioth�que ne contient aucun v�hicule personnalis�"
Const IDm_NomBibvehReserv� = GIRATIONVEHCOURT & " est un nom de biblioth�que r�serv� par GIRATION"

Private txtFichierAValider As Boolean

Private Sub cmdCancel_Click()
  Unload Me
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
' Cr�ation de la biblioth�que � partir des v�hicules s�lectionn�s
'**********************************************************************
Private Sub cmdOK_Click()
Dim numFich As Integer
Dim i As Integer
Dim VehicTab As StructV�hicule
Dim Abandon As Boolean
Dim numEnreg As Integer

    cmdParcourir = True

    If txtFichier <> "" Then
      numFich = FreeFile
      If ExistFich(txtFichier) Then Kill txtFichier
      Open txtFichier For Random As numFich Len = Len(VehicTab)
      numEnreg = 1
      Put #numFich, numEnreg, vehicVersion
      For i = 0 To lstV�hicules.ListCount - 1
        If lstV�hicules.Selected(i) Then
            'Mise au format StructV�hicule l'objet V�hicule
          ConvObjetStruct ColV�hicules(UCase(lstV�hicules.List(i))), VehicTab
            'Ecriture du v�hicule dans le fichier
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
      If FichierProt�g�(.FileName) Then Exit Sub
      txtFichier = .FileName
      If StrComp(nomCourt(txtFichier, SansExtension:=False), GIRATIONVEHCOURT, vbTextCompare) = 0 Then
        MsgBox IDm_NomBibvehReserv�
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
  If lstV�hicules.ListCount = 0 Then
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

    'Affichage centr� de la fen�tre
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
  
  ' Aide contextuelle
    Me.HelpContextID = EXPORTVEHICUL
    
  For Each Vehic In ColV�hicules
    'Cr�ation de la table de correspondance entre le num�ro de v�hicule dans la collection et l'item dans la listbox
    If Not Vehic.Prot�g� Then
    ' Seuls les v�hicules personnalis�s sont exportables
      lstV�hicules.AddItem Vehic.Nom
    End If
  Next
  lstV�hicules.ListIndex = -1
End Sub

'*****************************************
' L'utilisateur a coch� ou d�coch� un item
'******************************************
Private Sub lstV�hicules_ItemCheck(Item As Integer)
  cmdOK.Enabled = (lstV�hicules.SelCount <> 0)
End Sub


Private Sub txtFichier_Change()
  txtFichierAValider = True
End Sub

'*****************************************
' Validation du nom de fichier � cr�er
'******************************************
Private Sub txtFichier_Validate(Cancel As Boolean)
  If txtFichier = "" Then Exit Sub
  
  If Extension(txtFichier) = "" Then txtFichier = txtFichier + dlgFichier.DefaultExt
  
  If StrComp(nomCourt(txtFichier, SansExtension:=False), GIRATIONVEHCOURT, vbTextCompare) = 0 Then
    MsgBox IDm_NomBibvehReserv�
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

