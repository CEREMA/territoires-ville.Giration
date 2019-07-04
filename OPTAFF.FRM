VERSION 5.00
Begin VB.Form frmOptAff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options d'affichage"
   ClientHeight    =   6345
   ClientLeft      =   1380
   ClientTop       =   1110
   ClientWidth     =   6795
   FillColor       =   &H000000FF&
   Icon            =   "OPTAFF.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   320
      Left            =   5520
      TabIndex        =   19
      Top             =   5880
      Width           =   1092
   End
   Begin VB.CheckBox chkSauver 
      Caption         =   "Conserver comme valeurs par défaut"
      Height          =   372
      Left            =   3240
      TabIndex        =   16
      Top             =   5040
      Width           =   3132
   End
   Begin VB.Frame fraVéhicule 
      Caption         =   "Représentation du véhicule"
      Height          =   2295
      Left            =   3240
      TabIndex        =   32
      Top             =   2400
      Width           =   3375
      Begin VB.CheckBox chkMarkPosition 
         Caption         =   "Indiquer la position courante"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtPas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Text            =   " 5"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkMarkSens 
         Caption         =   "Marquer le sens de circulation"
         Height          =   372
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2412
      End
      Begin VB.TextBox txtPas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1080
         ScrollBars      =   1  'Horizontal
         TabIndex        =   12
         Text            =   " 5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSens 
         Caption         =   "positions"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   36
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblSens 
         Caption         =   "Toutes les "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblPas 
         Caption         =   "m"
         Height          =   252
         Index           =   1
         Left            =   1680
         TabIndex        =   34
         Top             =   360
         Width           =   252
      End
      Begin VB.Label lblPas 
         Caption         =   "Tous les "
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.Frame fraGeneral 
      Height          =   1935
      Left            =   3240
      TabIndex        =   29
      Top             =   240
      Width           =   3375
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   7
         Left            =   2880
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.OptionButton optCoulFDP 
         Caption         =   "Couleur unique"
         Enabled         =   0   'False
         Height          =   252
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optCoulFDP 
         Caption         =   "Couleurs d'origine"
         Enabled         =   0   'False
         Height          =   252
         Index           =   0
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   8
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.CheckBox chkGraduer 
         Caption         =   "Graduations"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1212
      End
      Begin VB.CheckBox chkFDP 
         Caption         =   "Fond de plan"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      CausesValidation=   0   'False
      Height          =   320
      Left            =   4320
      TabIndex        =   18
      Top             =   5880
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   3120
      TabIndex        =   17
      Top             =   5880
      Width           =   1092
   End
   Begin VB.Frame fraRemorque2 
      Caption         =   "Remorque 2"
      Height          =   1452
      Left            =   240
      TabIndex        =   26
      Top             =   4080
      Width           =   2412
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   6
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   2
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.CheckBox chkRouArr 
         Caption         =   "roues"
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1212
      End
      Begin VB.CheckBox chkCarross 
         Caption         =   "carrosserie"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Frame fraRemorque1 
      Caption         =   "Remorque 1"
      Height          =   1452
      Left            =   240
      TabIndex        =   23
      Top             =   2400
      Width           =   2412
      Begin VB.CheckBox chkCarross 
         Caption         =   "carrosserie"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1212
      End
      Begin VB.CheckBox chkRouArr 
         Caption         =   "roues"
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1212
      End
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   1
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   5
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
   End
   Begin VB.Frame fraTracteur 
      Caption         =   "Tracteur"
      Height          =   1932
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2412
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   4
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   3
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox Coul 
         BackColor       =   &H000000FF&
         Height          =   252
         Index           =   0
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.CheckBox chkRouArr 
         Caption         =   "roues arrière"
         Height          =   372
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1212
      End
      Begin VB.CheckBox chkRouAvt 
         Caption         =   "roues avant"
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1212
      End
      Begin VB.CheckBox chkCarross 
         Caption         =   "carrosserie"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1212
      End
   End
End
Attribute VB_Name = "frmOptAff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmOptAff   -   OPTAFF.FRM
'   Feuille permettant de spécifier les options d'affichage de la trajectoire

'**************************************************************************************
Option Explicit

Private sauvPas(1) As String

' Le tableau des couleurs est indicé comme suit
' 0 à 2 : carrosseries
' 3 : roues avant
' 4 à 6 : roues arrière
' 7 : fond de plan
' 8 : graduations


Private Sub chkCarross_Click(Index As Integer)

  Coul(Index).Visible = (chkCarross(Index) = vbChecked)
  
End Sub

' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la focntion KeyPress semble en mesure de réafficher la virgule
Private Sub txtPas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True
'    If flagVirgule Then alertVirgule = True
'    If flagVirgule Then KeyCode = 188
'  End If

End Sub

Private Sub txtPas_KeyPress(Index As Integer, KeyAscii As Integer)
  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False
End Sub


Private Sub chkRouAvt_Click()
    Coul(3).Visible = (chkRouAvt = vbChecked)
End Sub

Private Sub chkRouArr_Click(Index As Integer)
    
  Coul(Index + 4).Visible = (chkRouArr(Index) = vbChecked)

End Sub

Private Sub chkFDP_Click()
Dim i As Integer

  For i = 0 To 1
    If chkFDP = vbUnchecked Then
      optCoulFDP(i).Enabled = False
      Coul(7).Enabled = False
    Else
      optCoulFDP(i).Enabled = True
      Coul(7).Enabled = True
    End If
  Next
  
 End Sub

Private Sub optCoulFDP_Click(Index As Integer)
  If Index = 0 Then Coul(7).Visible = False Else Coul(7).Visible = True
End Sub

Private Sub chkGraduer_Click()
  If chkGraduer = vbUnchecked Then Coul(8).Visible = False Else Coul(8).Visible = True
End Sub

Private Sub chkMarkSens_Click()
    
'    lblSens(0).Visible = chkMarkSens <> vbUnchecked
'    lblSens(1).Visible = chkMarkSens <> vbUnchecked
'    txtPas(1).Visible = chkMarkSens <> vbUnchecked
    txtPas(1).Enabled = chkMarkSens <> vbUnchecked

End Sub

Private Sub Coul_Click(Index As Integer)
  
  On Error Resume Next
  With frmColor
    .color = Coul(Index).BackColor
    .Show vbModal
    Coul(Index).BackColor = .color
  End With
  
End Sub

Private Sub Form_Load()
Dim i As Integer

      If FichierJournal Then Write #numFichLog, "Ouverture Options d'affichage"
    
'      DégraisserFonte Me
    
    'Affichage centré de la fenêtre
      Me.ScaleMode = 1
      Me.Left = (Screen.Width - Me.Width) / 2
      Me.Top = (Screen.Height - Me.Height) / 2
  
  ' Aide contextuelle
  Me.HelpContextID = OPTAFFICHAGE
  
 
  With OptionAffichage(NumFille)
    For i = 0 To 2
      If .Carross(i).ADessiner = True Then chkCarross(i) = vbChecked
      Coul(i).BackColor = .Carross(i).couleur
      If .RouArr(i).ADessiner = True Then chkRouArr(i) = 1
      Coul(i + 4).BackColor = .RouArr(i).couleur
    Next
    If .RouAvt.ADessiner = True Then chkRouAvt = vbChecked
    Coul(3).BackColor = .RouAvt.couleur
    If .FDP.ADessiner Then chkFDP = vbChecked
    If .FDP.couleur = -1 Then
      optCoulFDP(0) = True
      Coul(7).BackColor = RGB(0, 0, 0)
    Else
      optCoulFDP(1) = True
      Coul(7).BackColor = .FDP.couleur
    End If
    If .Graduations.ADessiner Then chkGraduer = vbChecked
    Coul(8).BackColor = .Graduations.couleur
    
    If .sens Then chkMarkSens = vbChecked Else txtPas(1).Enabled = False
    txtPas(0) = Arrond(CDbl(.pas(0)), 2)
    txtPas(1) = CInt(.pas(1))
    
    If .position Then chkMarkPosition = vbChecked
    If flagImpr Then chkMarkPosition.Enabled = False
  End With
    
  If fCourante.nbRemorque <> 2 Then
    chkCarross(2).Enabled = False
    chkRouArr(2).Enabled = False
    If fCourante.nbRemorque <> 1 Then
      chkCarross(1).Enabled = False
      chkRouArr(1).Enabled = False
    End If
  End If
    
  If fCourante.FDP = "" Then
      chkFDP.Enabled = False
      If chkFDP = vbChecked Then chkFDP = vbGrayed
      optCoulFDP(0).Enabled = False
      optCoulFDP(1).Enabled = False
  End If

End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

Private Sub cmdOK_Click()
Dim i As Integer

' L'utilisateur a fait RETOUR-CHARIOT
''If ctrlRC(Me, "pas") Then Exit Sub

With OptionAffichage(NumFille)
  For i = 0 To fCourante.nbRemorque
    With .Carross(i)
      If chkCarross(i) = vbUnchecked Then .ADessiner = False Else .ADessiner = True: .couleur = Coul(i).BackColor: fCourante.colPlans(GirPlanReserve(i)).couleur = .couleur
    End With
    With .RouArr(i)
      If chkRouArr(i) = vbUnchecked Then .ADessiner = False Else .ADessiner = True: .couleur = Coul(i + 4).BackColor: fCourante.colPlans(GirPlanReserve(i + 4)).couleur = .couleur
    End With
  Next
  With .RouAvt
      If chkRouAvt = vbUnchecked Then .ADessiner = False Else .ADessiner = True: .couleur = Coul(3).BackColor: fCourante.colPlans(GirPlanReserve(3)).couleur = .couleur
  End With
  With .FDP
      If chkFDP = vbUnchecked Then
        .ADessiner = False
      Else
        .ADessiner = True
        If optCoulFDP(0) Then .couleur = -1 Else .couleur = Coul(7).BackColor
      End If
  End With
  With .Graduations
      If chkGraduer = vbUnchecked Then .ADessiner = False Else .ADessiner = True: .couleur = Coul(8).BackColor
  End With
  .pas(0) = txtPas(0)
  If chkMarkSens = vbUnchecked Then
    .sens = False
  Else
    .sens = True
    .pas(1) = txtPas(1)
  End If
  .position = (chkMarkPosition = vbChecked)
End With    ' OPtionAffichage(Numfille)


If chkSauver Then
' écrire dans la base de registres
  EcrireRegistry
  DefOptAff = OptionAffichage(NumFille)
End If
Me.Hide


If flagImpr Then
  ' Appel depuis la feuille d'impression : suppression des caractéristiques d'échelle qui y ont été créées
  fCourante.Milieux.Remove fCourante.Milieux.Count
  fCourante.Echelles.Remove fCourante.Echelles.Count
End If

      If FichierJournal Then Write #numFichLog, "Options d'affichage OK"

fCourante.flagRecalcul = True
'v3.3 : Mise en commentaire(02/05/07)
'ne sert à rien et en + par effet de bord peut générer l'erreur 5
'fCourante.Form_Activate
'mais remplacé par ce qui sert vraiment(15/06/07)
fCourante.AffecterOptAffichage

'frmAnnulImpr.Show
   
On Error GoTo GestErr

'frmAnnulImpr.lblEnCours = "Rafraichissement en cours..."
Rafraichir

If flagImpr Then
' Remise en condition de frmImprime comme lors de son 1er chargement
  flagImpr = False
  Set ObjetDessin = Printer
  frmImprim.EchelAjustee
End If
  
'If Not frmAnnulImpr.Annul Then Unload Me
Unload Me
'Unload frmAnnulImpr

Exit Sub

GestErr:
  If Err = 102 Then     ' Annulation du rafraichissement demandée
    Me.Visible = True ' Réaffichage de la feuille pour rechoisir les options
    cmdCancel.Enabled = False
    Resume Next
  End If
  
End Sub

Private Sub cmdCancel_Click()

      If FichierJournal Then Write #numFichLog, "Options d'affichage Annuler"
      Unload Me
      
End Sub


Private Sub EcrireRegistry()
' Mise à jour de la base de registres
  Dim res%
  
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Tracteur", Setting:=chkCarross(0)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Remorque1", Setting:=chkCarross(1)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Remorque2", Setting:=chkCarross(2)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues avant", Setting:=chkRouAvt
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues arrière", Setting:=chkRouArr(0)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues remorque1", Setting:=chkRouArr(1)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Roues remorque2", Setting:=chkRouArr(2)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Fond de plan", Setting:=chkFDP
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Graduations", Setting:=chkGraduer
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Marquage sens", Setting:=chkMarkSens
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Indication position", Setting:=chkMarkPosition
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Increment1", Setting:=CSng(txtPas(0))
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Increment2", Setting:=CSng(txtPas(1))
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur1", Setting:=decompRGB(Coul(0).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur2", Setting:=decompRGB(Coul(1).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur3", Setting:=decompRGB(Coul(2).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur4", Setting:=decompRGB(Coul(3).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur5", Setting:=decompRGB(Coul(4).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur6", Setting:=decompRGB(Coul(5).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur7", Setting:=decompRGB(Coul(6).BackColor)
  If OptionAffichage(NumFille).FDP.couleur = -1 Then
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur8", Setting:="-1"
  Else
    SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur8", Setting:=decompRGB(Coul(7).BackColor)
  End If
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur7", Setting:=decompRGB(Coul(6).BackColor)
  SaveSetting Appname:=App.Title, SECTION:="Affichage", Key:="Couleur9", Setting:=decompRGB(Coul(8).BackColor)

End Sub

Private Sub txtPas_GotFocus(Index As Integer)
  sauvPas(Index) = txtPas(Index)
End Sub

Private Sub txtPas_Validate(Index As Integer, Cancel As Boolean)

  If Trim(txtPas(Index)) = "" Then
    MsgBox ID_Saisie & " " & Idm_Obligatoire
    Cancel = True
  ElseIf controleNumeric(txtPas(Index), StrictPositif:=True) Then
    Cancel = True
  End If
  
  If Cancel Then
    txtPas(Index) = sauvPas(Index)
  Else
    If Index = 0 Then   ' arrondi aux 5cm supérieurs: Cahierdes Charges §4 - p 14
      txtPas(Index) = Arrond(CDbl(txtPas(Index)), 2, 5)
    Else
      txtPas(Index) = Arrond(CDbl(txtPas(Index)), 0)
    End If
  End If

End Sub
