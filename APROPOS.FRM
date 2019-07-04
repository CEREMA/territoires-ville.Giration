VERSION 5.00
Begin VB.Form frmApropos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A propos de Giration"
   ClientHeight    =   2970
   ClientLeft      =   1815
   ClientTop       =   2505
   ClientWidth     =   4200
   Icon            =   "APROPOS.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   2520
      Width           =   1332
   End
   Begin VB.Image imgLogo 
      Height          =   1215
      Index           =   1
      Left            =   480
      Picture         =   "APROPOS.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image imgLogo 
      Height          =   1215
      Index           =   0
      Left            =   2400
      Picture         =   "APROPOS.frx":3FC1E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      Caption         =   "Giration Version "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "frmApropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmApropos    -   Fichier APROPOS.FRM

'**************************************************************************************
Option Explicit

Private Sub cmdOK_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  Set Icon = MDIGiration.Icon
  
  cmdOK.Font.Bold = False
  
  'Affichage centré de la fenêtre
  Me.ScaleMode = 1
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  
  lblVersion.Caption = lblVersion.Caption & GirationVersion & vbCrLf
  If VersionDemo Then
    lblVersion.Caption = lblVersion.Caption & LBLICENCE & NumeroLicence
  Else
    'lblVersion.Caption = lblVersion.Caption ' & ID_Licence & " " & SerialNumber
    lblVersion.Caption = lblVersion.Caption & ID_Licence & " " & App.Revision & vbCrLf & LBLICENCE & NumeroLicence

  End If
    
End Sub


