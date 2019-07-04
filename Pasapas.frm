VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmPas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pas à pas"
   ClientHeight    =   1590
   ClientLeft      =   2775
   ClientTop       =   5010
   ClientWidth     =   4245
   Icon            =   "Pasapas.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4245
   Begin VB.Timer tmrPasAPas 
      Enabled         =   0   'False
      Left            =   240
      Top             =   1080
   End
   Begin VB.Frame fraVitesse 
      Caption         =   "Vitesse de défilement"
      Height          =   852
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3975
      Begin VB.CheckBox chkManuel 
         Caption         =   "Manuel"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar hsbVitesse 
         Height          =   252
         Left            =   360
         Max             =   100
         Min             =   1
         TabIndex        =   3
         Top             =   480
         Value           =   30
         Width           =   2172
      End
      Begin VB.Label lblRapide 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblLent 
         Caption         =   "Lent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "&Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   972
   End
   Begin MCI.MMControl mciMagneto 
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   529
      _Version        =   393216
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      EjectEnabled    =   -1  'True
      Silent          =   -1  'True
      AutoEnable      =   0   'False
      PlayVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "frmPas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmPas   -   PASAPAS.FRM
'   Feuille visualisant le véhicule pas à pas

'**************************************************************************************
Option Explicit

Private numpos As Integer, numposprec As Integer, avancement As Integer

Private Sub chkManuel_Click()
' Ajout v3.0.206
  With mciMagneto
    If chkManuel = vbChecked Then
      .StopEnabled = False
      .NextEnabled = numpos < fCourante.maxPosition Or (numpos = -1 And numposprec < fCourante.maxPosition)
      .PrevEnabled = numpos >= 0 Or (numpos = -1 And numposprec > 0)
    Else
      If numpos <> -1 Then .StopEnabled = True
      .NextEnabled = True
      .PrevEnabled = True
    End If
  End With
  ActiveScroll (chkManuel = vbChecked)
End Sub

Private Sub ActiveScroll(ByVal Manuel As Boolean)
  lblLent.Enabled = Not Manuel
  lblRapide.Enabled = Not Manuel
  hsbVitesse.Enabled = Not Manuel
End Sub

Private Sub cmdFermer_Click()
Dim i As Integer

  Me.Hide
      If FichierJournal Then Write #numFichLog, "Fin du pas à pas"
  gtEffacement = True
  fCourante.desPosition numposprec
  gtEffacement = False
  fCourante.dessiner "TOUT"
  fCourante.VitPas = hsbVitesse
  fCourante.PasManuel = (chkManuel = vbChecked)
  Unload Me
  
  With MDIGiration
  ' Mise en commentaire AV v3.2.1 : 04/04/2000 - On préfère ignorer les actions que mettre en grisé
'    .Enabled = True
'    .mnuBarre(0).Enabled = True
'    .mnuBarre(2).Enabled = True
'    .tbrGiration.Enabled = True
  End With
  With fCourante
    .picBoutons.Enabled = True
'''    .cmdBoutonOrigine.Enabled = True
'    .cmdVéhicule.Enabled = True
    .Grille.Enabled = True
  End With
  GriserMenus Etat:=True

      If FichierJournal Then Write #numFichLog, "Fin du pas à pas OK"
  
End Sub

Private Sub Form_Load()
    Dim i As Integer
  Dim Button  As MSComctlLib.Button
  
    If FichierJournal Then Write #numFichLog, "Ouverture du pas à pas"
  
  Set Icon = fCourante.Icon
  numposprec = 0
  numpos = -1
  fCourante.linMarque(0).Visible = False
  fCourante.linMarque(1).Visible = False
  
  With MDIGiration
    Dim ItemMenu As Control
    For Each ItemMenu In .mnuOutils 'Activation du menu Outils : Cadrage et Rafraichir
      If ItemMenu.Caption <> "-" And ItemMenu.Index <> MNUCADRAGE And ItemMenu.Index <> MNURAFRAICHIR Then ItemMenu.Enabled = False
    Next
  
  ' Mise en commentaire AV v3.2.1 : 04/04/2000 - On préfère ignorer les actions que mettre en grisé
'    .mnuBarre(0).Enabled = False
'    .mnuBarre(2).Enabled = False
'    .tbrGiration.Enabled = False
  End With
  With fCourante
    .picBoutons.Enabled = False
'''    .cmdBoutonOrigine.Enabled = False
'    .cmdVéhicule.Enabled = False
    .Grille.Enabled = False
  End With
    
    fCourante.effTout
    Move fCourante.Grille.Left + fCourante.Grille.Width - Width + 550, fCourante.Top + fCourante.Grille.Top + 1250
    
    ' Aide contextuelle
    Me.HelpContextID = PASAPAS
    
    If fCourante.VitPas = 0 Then
      hsbVitesse = 30
    Else
      hsbVitesse = fCourante.VitPas
    End If
    If fCourante.PasManuel Then chkManuel = vbChecked
    

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = vbFormControlMenu Then
  ' Fermeture au moyen du bouton système
    cmdFermer = True
  End If
  
End Sub

Private Sub hsbVitesse_Change()
' vitesse=1 correspond à 1 toutes les 1 s       (1000 ms)
' vitesse=100 correspond à 1 tous tes 1/100s (10 ms)
' v3.0.205 : 5000 au lieu de 1000 (sinon trop vite par défaut)

  tmrPasAPas.Interval = 5000 / hsbVitesse
  DoEvents
  
End Sub

Private Sub mciMagneto_BackClick(Cancel As Integer)

ActivBoutonMagneto Marche:=True

If numpos = -1 Then
  numpos = Min(numposprec, fCourante.maxPosition) ' Le Bouton STOP a été activé - redémarrage à la fin ou au même point
  numpos = fCourante.maxPosition ' v3.0.205
ElseIf avancement = 1 Then  ' Changement de sens de parcours
  numpos = numpos - 1
End If
avancement = -1
tmrPasAPas.Enabled = True
dessiner

End Sub

Private Sub mciMagneto_EjectClick(Cancel As Integer)

  cmdFermer = True
  
End Sub

Private Sub mciMagneto_NextClick(Cancel As Integer)
  
  If chkManuel = vbUnchecked Then
    mciMagneto_StepClick False
    Exit Sub
  End If
  
  mciMagneto.PrevEnabled = True
  mciMagneto.BackEnabled = True

  If avancement = -1 Then
    numpos = numposprec + 1 ' Changement de sens de parcours
  ElseIf numpos = -1 Then
    numpos = numposprec + 1
  End If
  avancement = 1
  dessiner
  
  If numpos > fCourante.maxPosition Then mciMagneto.StepEnabled = False: mciMagneto.NextEnabled = False
  
End Sub

Private Sub mciMagneto_PauseClick(Cancel As Integer)

  tmrPasAPas.Enabled = False
  
  ActivBoutonMagneto Marche:=False, Pause:=True
End Sub

Private Sub mciMagneto_PlayClick(Cancel As Integer)

  tmrPasAPas.Enabled = True
  
  ActivBoutonMagneto Marche:=True
  
  ' v3.0.0
  If numpos = -1 Then
    numpos = 0 ' Le Bouton STOP a été activé - redémarrage au départ
  ElseIf avancement = -1 Then ' Changement de sens de parcours
    numpos = numpos + 1
  End If
  
  ' v3.0.205
  If numpos = -1 Then ' Le Bouton STOP a été activé
    If avancement = 1 Then numpos = 0 Else numpos = fCourante.maxPosition
  End If
  
  avancement = 1
  dessiner
  
End Sub

Private Sub mciMagneto_PrevClick(Cancel As Integer)
  
  If chkManuel = vbUnchecked Then
    mciMagneto_BackClick False
    Exit Sub
  End If
  
  mciMagneto.NextEnabled = True
  mciMagneto.StepEnabled = True

  If avancement = 1 Then
    numpos = numposprec - 1 ' Changement de sens de parcours
  ElseIf numpos = -1 Then
    numpos = numposprec - 1
  End If
  avancement = -1
  dessiner
  
  If numpos < 0 Then mciMagneto.BackEnabled = False: mciMagneto.PrevEnabled = False

End Sub

' Evènement inutilisé en v3.0.0
Private Sub mciMagneto_StepClick(Cancel As Integer)

' tmrPasAPas.Enabled = True v3.0.0
' numpos = fCourante.maxPosition    v3.0.0

  tmrPasAPas.Enabled = True
  
  ActivBoutonMagneto Marche:=True
  
  If numpos = -1 Then
    numpos = 0 ' Le Bouton STOP a été activé - redémarrage au départ
  ElseIf avancement = -1 Then ' Changement de sens de parcours
    numpos = numpos + 1
  End If

  avancement = 1
  dessiner

End Sub

Private Sub mciMagneto_StopClick(Cancel As Integer)
  
  tmrPasAPas.Enabled = False
  ActivBoutonMagneto Marche:=False, Pause:=False
'  mciMagneto.NextEnabled = numpos < fCourante.maxPosition  Mise en commentaire : v 3.0.206
'  mciMagneto.PrevEnabled = numpos > 0                id
  
' Indicateur permettant de distinguer le bouton Stop de Pause
' Après Pause, on repart du point courant - Après STOP, on repart du point courant si Marche arrière, sinon on repart au début
  numpos = -1

End Sub

Private Sub tmrPasAPas_Timer()

  If numpos > fCourante.maxPosition Or numpos < 0 Then
    mciMagneto_StopClick Cancel:=False
    Exit Sub
  End If
  
  dessiner
End Sub

Private Sub ActivBoutonMagneto(ByVal Marche As Boolean, Optional ByVal Pause As Boolean)
  With mciMagneto
    chkManuel.Enabled = Not Marche
    .EjectEnabled = Not Marche
    .PauseEnabled = Marche
'    .PlayEnabled = Not Marche Mise en commentaire : v3.0.205
'    .BackEnabled = Not Marche  Mise en commentaire : v3.0.205
    If Not Pause Then
      .StopEnabled = Marche
'      .StepEnabled = Not Marche   ' v3.0.205
'      .BackEnabled = Not Marche   ' v3.0.205
    End If
    .StepEnabled = Not Marche   ' v3.0.206
    .BackEnabled = Not Marche   ' v3.0.206
    If chkManuel = vbChecked Then
      .NextEnabled = numpos < fCourante.maxPosition And Not Marche    ' v3.0.205
      .PrevEnabled = numpos > 0 And Not Marche                ' v3.0.205
    Else
      .NextEnabled = .StepEnabled   ' v3.0.206
      .PrevEnabled = .BackEnabled   ' v3.0.206
    End If
  End With
  
'  With fCourante
'    .cmdZoom.Enabled = Not Marche
'    .cmdZAvant.Enabled = Not Marche
'    .cmdPAN.Enabled = Not Marche
'  End With
'  MDIGiration.mnuBarre(1).Enabled = Not Marche
  MDIGiration.Enabled = Not Marche
  
End Sub

Public Sub dessiner(Optional ByVal RetourZoom As Boolean)

  If RetourZoom Then  ' Ajout AV 08.02.2000 : Possibilité des Fonctions de Zoom avec le pas à pas
    fCourante.desPosition numposprec
    Exit Sub
  End If

  gtEffacement = True
  fCourante.desPosition numposprec
  gtEffacement = False

'  fCourante.picTrajectoire.CurrentX = 500
'  fCourante.picTrajectoire.CurrentY = 500
'  fCourante.picTrajectoire.Print Format(numpos, "00")

  fCourante.desPosition numpos
  numposprec = numpos
  numpos = numpos + avancement

End Sub

