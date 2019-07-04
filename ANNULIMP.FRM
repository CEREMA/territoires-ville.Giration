VERSION 5.00
Begin VB.Form frmAnnulImpr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GIRATION"
   ClientHeight    =   1485
   ClientLeft      =   2310
   ClientTop       =   5865
   ClientWidth     =   3270
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label lblEncours 
      Alignment       =   2  'Center
      Caption         =   "Impression en cours ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmAnnulImpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmAnnulImpr     -  Fichier ANNULIMP.FRM
'   Interruption possible par l'utilisateur de l'Impression

'**************************************************************************************

Option Explicit

Public Annul As Boolean

Private Sub cmdCancel_Click()

  Annul = True
  DoEvents
  
End Sub


Private Sub Form_Activate()

On Error Resume Next

cmdCancel.SetFocus
 
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
