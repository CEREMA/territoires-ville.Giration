VERSION 5.00
Begin VB.Form frmDéfileDXF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import de fichier DXF"
   ClientHeight    =   1485
   ClientLeft      =   1770
   ClientTop       =   5895
   ClientWidth     =   4140
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPanneau 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.Timer tmrDéfile 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   840
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   372
      Left            =   1284
      TabIndex        =   0
      Top             =   1008
      Width           =   1572
   End
   Begin VB.Label lblPourCent 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmDéfileDXF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmDéfileDXF    -   Fichier DEFILDXF.FRM
'   Fait défiler le nom du fichier DXF en cours de lecture et affiche le pourcentage d'avancement
'   Interruption possible par l'utilisateur de la lecture d'un fichier DXF

'**************************************************************************************

Option Explicit
  

Public Annul As Boolean

Private Sub cmdCancel_Click()

  Annul = True
  
End Sub


Private Sub Form_Load()

  'cmdCancel.Font.Bold = False
  
'Affichage centré de la fenêtre
  'Me.ScaleMode = 1
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2

  Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  ' On interdit à l'utilisateur de fermer la fenêtre
  
  If UnloadMode = vbFormControlMenu Then Cancel = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Annul = False
  
End Sub

Private Sub tmrDéfile_Timer()
' A chaque interruption du Timer :
'   renvoi en dernière position du premier caractère de panneau,donnant l'impression d'un défilement
Dim prem As String, suite As String
  
  prem = Left(txtPanneau, 1)
  ' suppression du 1er caractère
  suite = Right(txtPanneau, Len(txtPanneau) - 1)
  ' rajout du premier caractère à la fin de la chaine
  txtPanneau = suite + prem
  
End Sub
