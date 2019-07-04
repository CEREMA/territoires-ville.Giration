VERSION 5.00
Begin VB.Form frmNewVehicule 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nouveau véhicule"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PICbOUTONS 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2070
      Left            =   3270
      ScaleHeight     =   2070
      ScaleWidth      =   1335
      TabIndex        =   6
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Aide"
         Height          =   320
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuler"
         CausesValidation=   0   'False
         Height          =   320
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   320
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Frame fraVehicule 
      Caption         =   " Véhicule "
      Height          =   1812
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3132
      Begin VB.TextBox txtNom 
         Height          =   285
         Left            =   720
         MaxLength       =   20
         TabIndex        =   0
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Bi-articulé"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Articulé"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Simple"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblNom 
         Caption         =   "Nom :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmNewVehicule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Aout 2006

'   Réalisation : André VIGNAUD

'   Module de feuille : frmNewVéhicule   -   NewVehicule.FRM
'   Saisie du type et du nom d'un nouveau véhicule

'**************************************************************************************
Option Explicit
Public ValidOK As Boolean

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

Private Sub cmdOK_Click()
  If ExistVeh(txtNom.Text, ColVéhicules) Then
    MsgBox "Ce véhicule existe déjà dans la bibliothèque"
  Else
    ValidOK = True
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
   ' Aide en ligne contexte
   Me.HelpContextID = NEWVEHICUL

End Sub

Private Sub optTypVeh_Click(Index As Integer)
  If Len(txtNom.Text) > 0 Then
    cmdOK.Enabled = True
  End If
End Sub

Private Sub txtNom_Change()
  If Len(txtNom.Text) = 0 Then
    cmdOK.Enabled = False
  ElseIf Numopt(optTypVeh) <> -1 Then
    cmdOK.Enabled = True
  End If
End Sub
