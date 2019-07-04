VERSION 5.00
Begin VB.Form frmColor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix d'une couleur"
   ClientHeight    =   3135
   ClientLeft      =   1665
   ClientTop       =   5205
   ClientWidth     =   4035
   Icon            =   "COLOR.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picActiveCouleur 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   1920
      Width           =   372
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   320
      Left            =   2040
      TabIndex        =   18
      Top             =   2640
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   600
      TabIndex        =   17
      Top             =   2640
      Width           =   1212
   End
   Begin VB.PictureBox picChariotCouleur 
      Enabled         =   0   'False
      Height          =   252
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   2160
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   9
      Left            =   1440
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   15
      Top             =   960
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   15
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   14
      Top             =   1320
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   14
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   13
      Top             =   1320
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   13
      Left            =   1440
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   12
      Top             =   1320
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   12
      Left            =   840
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   11
      Top             =   1320
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   11
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   10
      Top             =   960
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   10
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   9
      Top             =   960
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   840
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   960
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   600
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   600
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   1440
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   600
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   840
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   600
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   240
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   240
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1440
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   240
      Width           =   372
   End
   Begin VB.PictureBox picCouleur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   840
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   240
      Width           =   372
   End
   Begin VB.Label lblCouleur 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Couleur :"
      Height          =   252
      Left            =   840
      TabIndex        =   20
      Top             =   1920
      Width           =   732
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmColor   -   COLOR.FRM
'   Feuille de choix de couleur appelée depuis frmOptAff (Options d'affichage)

'**************************************************************************************
Option Explicit
Dim hg(2), inc_h, inc_lg As Single
Public color As Long
Dim SauvColor As Long
Dim textCoul(15) As String

Private Sub cmdCancel_Click()
  
  color = SauvColor
  Unload Me

End Sub

Private Sub Form_Load()
Dim i%

Const ID_Noir = "Noir"
Const ID_Bleu = "Bleu"                             '       Bleu nuit
Const ID_Vert = "Vert"                              '       Vert cacadois
Const ID_Cyan = "Cyan"                            '       Bleu-Vert
Const ID_Rouge = "Rouge"                          '       Marron
Const ID_Magenta = "Magenta"                       '       Magenta-->Mauve
Const ID_Jaune = "Jaune"                           '       Jaune cacadois
Const ID_Blanc = "Blanc"                            '      Gris clair
Const ID_Gris = "Gris"                              '      Gris foncé
Const ID_BleuClair = "Bleu clair"                      '      Bleu
Const ID_VertClair = "Vert clair"                     '      Vert clair
Const ID_CyanClair = "Cyan clair"                   '       Cyan
Const ID_RougeClair = "Rouge clair"                 '      Rouge
Const ID_MagentaClair = "Magenta clair"              '     Rose
Const ID_JauneClair = "Jaune clair"                  '     Jaune
Const ID_BlancVif = "Blanc vif"                      '     Blanc

' DégraisserFonte Me
  
  textCoul(0) = ID_Noir                              '       Noir
  textCoul(1) = ID_Bleu                             '       Bleu nuit
  textCoul(2) = ID_Vert                              '       Vert cacadois
  textCoul(3) = ID_Cyan                            '       Bleu-Vert
  textCoul(4) = ID_Rouge                          '       Marron
  textCoul(5) = ID_Magenta                       '       Magenta-->Mauve
  textCoul(6) = ID_Jaune                           '       Jaune cacadois
  textCoul(7) = ID_Blanc                            '      Gris clair
  textCoul(8) = ID_Gris                              '      Gris foncé
  textCoul(9) = ID_BleuClair                      '      Bleu
  textCoul(10) = ID_VertClair                     '      Vert clair
  textCoul(11) = ID_CyanClair                   '       Cyan
  textCoul(12) = ID_RougeClair                 '      Rouge
  textCoul(13) = ID_MagentaClair              '     Rose
  textCoul(14) = ID_JauneClair                  '     Jaune
  textCoul(15) = ID_BlancVif                      '     Blanc

  
  hg(0) = picCouleur(0).Left
  inc_lg = picCouleur(1).Left - hg(0)
  hg(1) = picCouleur(0).Top
  inc_h = picCouleur(4).Top - hg(1)
  SauvColor = color
  
   For i = 0 To 15
    picCouleur(i).BackColor = QBColor(i)
    If QBColor(i) = color Then picCouleur_Click i
  Next
  
End Sub

Private Sub cmdOK_Click()
    
    Unload Me

End Sub

Private Sub picCouleur_Click(Index As Integer)
Dim numlig%, numcol%

  numcol = Index Mod 4
  numlig = (Index - numcol) / 4
  color = QBColor(Index)
  
  ' picChariotCouleur se déplace sur la bonne case couleur
  picChariotCouleur.Move hg(0) + numcol * inc_lg, hg(1) + numlig * inc_h
  picChariotCouleur.BackColor = color
  'picActiveCouleur est la couleur 'témoin'
  picActiveCouleur.BackColor = color

End Sub
