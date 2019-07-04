VERSION 5.00
Begin VB.Form frmGrandDessin 
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   870
   ClientTop       =   1530
   ClientWidth     =   5280
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTrajectoire 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3252
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   120
      Width           =   5052
   End
End
Attribute VB_Name = "frmGrandDessin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmGrandDessin   -   FRMGIGA.FRM
'   Feuille visualisant ponctuellement  la trajectoire sur tout l'ensemble  de l'écran

'**************************************************************************************
Option Explicit

Private basGauche As New PT, hautDroit As New PT

Private Sub Form_Activate()
  
  Rafraichir

End Sub

Private Sub Form_Load()

  Me.Move 0, 0, Screen.Width, Screen.Height
  picTrajectoire.Top = Me.Top
  picTrajectoire.Left = Me.Left
  picTrajectoire.Width = Me.Width
  picTrajectoire.Height = Me.Height
  
    basGauche.X = fCourante.pmin.X
    basGauche.Y = fCourante.pmin.Y
    hautDroit.X = fCourante.pmax.X
    hautDroit.Y = fCourante.pmax.Y

  ' Si la fenêtre frmTraject a été redimensionnée, le centre de la vue n'est recalculé qu'au 1er Rafraichir
  ' Il faut le vérifier (voir Form_Resize dans frmTraject)
  verifRetaillage
  
  Set ObjetDessin = picTrajectoire
  gtOrx = ObjetDessin.ScaleWidth / 2  ' paramètres d'origine et d'échelle du dessin
  gtOry = ObjetDessin.ScaleHeight / 2
  xMaxPicture = ObjetDessin.ScaleWidth
  yMaxPicture = ObjetDessin.ScaleHeight
  ParamEcranZoom
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  
     fCourante.pmin.X = basGauche.X
    fCourante.pmin.Y = basGauche.Y
    fCourante.pmax.X = hautDroit.X
    fCourante.pmax.Y = hautDroit.Y
  
  fCourante.Form_Activate
  
End Sub


Private Sub picTrajectoire_DblClick()

If Not MDIGiration.mnuBarre(0).Enabled Then Exit Sub
  
  Unload Me
  Set ObjetDessin = fCourante.picTrajectoire
  
End Sub


