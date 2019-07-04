VERSION 5.00
Begin VB.Form frmVéhicule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Véhicule"
   ClientHeight    =   6600
   ClientLeft      =   150
   ClientTop       =   2430
   ClientWidth     =   11730
   Icon            =   "VEHICULE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   320
      Left            =   10320
      TabIndex        =   50
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CheckBox chkVerrou 
      Caption         =   "Verrouiller le véhicule"
      Height          =   372
      Left            =   480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.CommandButton cmdImprim 
      Caption         =   "&Imprimer"
      Height          =   320
      Left            =   10320
      TabIndex        =   22
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Frame fraCarTrac 
      Caption         =   "Caractéristiques du tracteur "
      Height          =   1452
      Left            =   240
      TabIndex        =   37
      Top             =   4440
      Width           =   5532
      Begin VB.ComboBox cboRayon 
         Height          =   288
         ItemData        =   "VEHICULE.frx":030A
         Left            =   120
         List            =   "VEHICULE.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   960
         Width           =   3612
      End
      Begin VB.TextBox txtRTracteur 
         Height          =   288
         Left            =   4680
         TabIndex        =   19
         Top             =   960
         Width           =   612
      End
      Begin VB.TextBox txtaVehMax 
         Height          =   288
         Left            =   4680
         TabIndex        =   18
         Top             =   480
         Width           =   612
      End
      Begin VB.Label lblAVehMax 
         Caption         =   "Vitesse max. de braquage en "
         Height          =   252
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   2652
      End
      Begin VB.Label lblUnite 
         Height          =   252
         Left            =   2880
         TabIndex        =   38
         Top             =   480
         Width           =   1692
      End
   End
   Begin VB.Frame fraCarGeom 
      Height          =   3612
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   7932
      Begin VB.TextBox txtNom 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame fraTracteur 
         Caption         =   "Tracteur"
         Height          =   3132
         Left            =   3960
         TabIndex        =   42
         Top             =   360
         Width           =   1092
         Begin VB.TextBox txtLEar 
            Height          =   288
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   2280
            Width           =   700
         End
         Begin VB.TextBox txtPAFar 
            Height          =   288
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtLEav 
            Height          =   288
            Left            =   240
            TabIndex        =   4
            Top             =   1920
            Width           =   700
         End
         Begin VB.TextBox txtE 
            Height          =   288
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   840
            Width           =   700
         End
         Begin VB.TextBox txtL 
            Height          =   288
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtPAFav 
            Height          =   288
            Index           =   0
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   700
         End
         Begin VB.TextBox txtDpat 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   2640
            Width           =   700
         End
      End
      Begin VB.Frame fraRemorque2 
         Caption         =   "Remorque2"
         Height          =   3132
         Left            =   6600
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   1092
         Begin VB.TextBox txtPAFav 
            Height          =   288
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   700
         End
         Begin VB.TextBox txtL 
            Height          =   288
            Index           =   2
            Left            =   240
            TabIndex        =   16
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtE 
            Height          =   288
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   700
         End
         Begin VB.TextBox txtPAFar 
            Height          =   288
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtLEar 
            Height          =   288
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   2280
            Width           =   700
         End
      End
      Begin VB.Frame fraRemorque1 
         Caption         =   "Remorque"
         Height          =   3132
         Left            =   5280
         TabIndex        =   40
         Top             =   360
         Width           =   1092
         Begin VB.TextBox txtPAFav 
            Height          =   288
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   700
         End
         Begin VB.TextBox txtL 
            Height          =   288
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtE 
            Height          =   288
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   700
         End
         Begin VB.TextBox txtPAFar 
            Height          =   288
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtLEar 
            Height          =   288
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   2280
            Width           =   700
         End
         Begin VB.TextBox txtDpat 
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   700
         End
      End
      Begin VB.Label lblNom 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "                              "
         Height          =   192
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label lblDpat 
         Caption         =   "D : Distance attache/essieu arrière "
         Height          =   288
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Width           =   3372
      End
      Begin VB.Label lblLEar 
         Caption         =   "LEar : Largeur essieu arrière avec pneus"
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   3612
      End
      Begin VB.Label lblLEav 
         Caption         =   "LEav : Largeur essieu avant avec pneus"
         Height          =   312
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   3612
      End
      Begin VB.Label lblL 
         Caption         =   "L : Largeur hors tout"
         Height          =   288
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   1692
      End
      Begin VB.Label lblPAFar 
         Caption         =   "PAFar : Porte-à-faux arrière"
         Height          =   288
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   2172
      End
      Begin VB.Label lblE 
         Caption         =   "E : Empattement"
         Height          =   288
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label lblPAFav 
         Caption         =   "PAFav : Porte-à-faux avant"
         Height          =   288
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2172
      End
   End
   Begin VB.Frame fraCarVeh 
      Caption         =   "Caractéristiques du véhicule "
      Height          =   1452
      Left            =   6000
      TabIndex        =   24
      Top             =   4440
      Width           =   5292
      Begin VB.Label lblBraqReduit 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Braquage réduit"
         Height          =   192
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label lblRayVehic 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   4080
         TabIndex        =   48
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lblRayVehic 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   4080
         TabIndex        =   47
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblRayVehic 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   46
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lblRayVehic 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   45
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblEntreMurs 
         Caption         =   "entre murs"
         Height          =   288
         Left            =   2520
         TabIndex        =   35
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label lblRayExter 
         Caption         =   "Rayons extérieurs"
         Height          =   252
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1812
      End
      Begin VB.Label lblEntreTrottoirs 
         Caption         =   "entre trottoirs"
         Height          =   288
         Left            =   3960
         TabIndex        =   33
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lblRayInter 
         Caption         =   "Rayons intérieurs"
         Height          =   288
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1932
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   10320
      TabIndex        =   20
      Top             =   240
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      CausesValidation=   0   'False
      Height          =   320
      Left            =   10320
      TabIndex        =   21
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label lblMsgRExtEntreMurs 
      Caption         =   $"VEHICULE.frx":0398
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   360
      TabIndex        =   51
      Top             =   6000
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Image Illustration 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1815
      Index           =   3
      Left            =   8400
      Picture         =   "VEHICULE.frx":0443
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image Illustration 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1815
      Index           =   2
      Left            =   8400
      Picture         =   "VEHICULE.frx":291F
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image Illustration 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1815
      Index           =   1
      Left            =   8400
      Picture         =   "VEHICULE.frx":63B1
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image Illustration 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1812
      Index           =   0
      Left            =   8400
      Picture         =   "VEHICULE.frx":81CF
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   3132
   End
End
Attribute VB_Name = "frmVéhicule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmVéhicule   -   VEHICULE.FRM
'   Saisie et/ou Visualisation des caractéristiques d'un véhicule de la bibliothèque
' Cette feuille est chargée depuis frmBibVéhicule
' ou depuis frmTraject (dans ce dernier cas, on ne peut que visualiser ou faire des simulations en modifiant le rayon du tracteur
' La variable contextBibliVeh permet de savoir dans quel cas on est

'**************************************************************************************
Option Explicit

' Déclaration des constantes de chaine

Const ID_SurPlace = "Sur place"

Const ID_PAFar = "Porte-à-faux arrière"
Const ID_Tracteur = "Tracteur"

Const Idm_Incomplet = "Saisie incomplète"
Const IDm_PasDeMemo = "Le véhicule ne sera pas mémorisé"
Const IDm_TropFaible = "trop faible"
Const IDm_ValeurGrande = "Attention : valeur grande"
Const IDm_BraqTropFaible = "Angle de braquage des roues trop faible pour ce tracteur "
Const IDm_BraqTropFort = "Angle de braquage des roues trop grand pour ce tracteur"
Const IDm_AttacheEloignée = "L'attache ne doit pas être éloignée de plus de la moitié de l'empattement"
Const IDm_RemorqueRayon = "La remorque ne permet pas au tracteur d'atteindre le rayon saisi" & vbCrLf & "Cette valeur entraine la rotation de la remorque sur place"

Public contextBibliVeh As Boolean

Private sauvPAFav(2) As String, sauvE(2) As String, sauvPAFar(2) As String, sauvL(2) As String, sauvLEav As String, sauvLEar(2) As String, sauvDpat(1) As String, sauvRVehMin As String, sauvaVehMax As String
Private calcule(2) As Boolean, calculOk As Boolean, calculDéjaFait As Boolean
Private R(4, 2)   As Double
Private idxillustr As Integer
'Private nbRemorque As Integer
Private Creation As Boolean, BraqReduitSignalé As Boolean
Private VehicTab As StructVéhicule
Private Modif As Boolean
Private mVéhicule As VEHICULE

Private Property Get nbRemorque() As Integer
  nbRemorque = mVéhicule.nbRemorque
End Property

Private Sub CacheMoi()
  Me.Hide
  DoEvents
End Sub

Private Sub cmdCancel_Click()
Dim Cancel As Boolean

  If contextBibliVeh Then
    If Modif Then
      Cancel = MsgBox("Abandonner la saisie du véhicule " & Me.lblNom, vbYesNo + vbDefaultButton2) = vbNo
    End If
    If Not Cancel Then
      CacheMoi    ' la feuille sera déchargée par frmBibVéhicule
      Creation = False
      Modif = False
    End If
    
  Else
  ' Contexte trajectoire : cmdCancel est le bouton par défaut -     txtL'utilisateur a fait RETOUR-CHARIOT
    If ActiveControl.name <> "cmdCancel" Then
      SendKeys "{TAB}"    ' pour forcer la vérification et les calculs
      txtRTracteur.SetFocus
    Else
      Unload Me
    End If
  End If
  
End Sub

'******************************************************************************
' Activation de la feuille
'******************************************************************************
Private Sub Form_Activate()

If contextBibliVeh Then
  HelpContextID = DEFVEHICUL
Else
  HelpContextID = CARACVEHICUL
End If

End Sub

'******************************************************************************
' DésActivation de la feuille
'******************************************************************************
Private Sub Form_Deactivate()

  MDIGiration.HelpContextID = SOMMAIRE  '0

End Sub

'******************************************************************************
' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la fonction KeyPress semble en mesure de réafficher la virgule
'******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False

End Sub

'******************************************************************************
' Chargement de la feuille
'******************************************************************************
Private Sub Form_Load()

  BraqReduitSignalé = True
  SetDeviceIndependentWindow Me
  cboRayon.ListIndex = 0
  lblUnite = libUnite(OptGen.unite) & "/" & ID_seconde
      
  If Printers.Count = 0 Or VersionDemo Then cmdImprim.Enabled = False
      
  If contextBibliVeh Then
    ' Autorisé pour le CERTU avec l'option  /V CERTU au lancement
    If VerrouVéhicule Then chkVerrou.Visible = True
    
    Set mVéhicule = frmBibVéhicule.Véhicule
    With mVéhicule
      lblNom.Caption = .Nom
      txtNom.Text = .Nom
      If Not ExistVeh(.Nom, ColVéhicules) Then
        Creation = True
        BraqReduitSignalé = False
        Me.txtNom.Visible = True
        
      Else
        If .Protégé Then
          If .nbUtilisation > 0 Or Not VerrouVéhicule Then
            cmdCancel.Caption = ID_Fermer
          End If
        Else
          txtNom.Visible = True
        End If
        AfficherVéhicule mVéhicule
        If .nbUtilisation > 0 Then   'Véhicule ne cours d'utilisation ds 1 trajectoire
          chkVerrou.Visible = False
        End If
        If frmBibVéhicule.EnVisu Then
          ' Fonctionnalité dont dispose uniqt le CERTU
          chkVerrou = vbChecked  ' cette activation permet d'inactiver tous les controles sans qu'il deviennent grisés
        End If
      End If
    End With      ' frmBibVéhicule
  
  Else    ' contexte Trajectoire
    Set mVéhicule = fCourante.Vehic
    lblNom = mVéhicule.Nom
    cmdCancel.Caption = ID_Fermer
    AfficherVéhicule mVéhicule
    chkVerrou = vbChecked  ' cette activation permet d'inactiver tous les controles sans qu'il deviennent grisés
    txtRTracteur.Enabled = True
  End If
  
  'Initialisation du titre de la fenêtre en fonction du choix de véhicule
  Select Case nbRemorque
  Case 0
      Caption = Majus(ID_Véhicule) & " " & ID_Simple
      fraTracteur.Caption = ""
      fraRemorque1.Visible = False
      fraRemorque2.Visible = False
      txtDpat(0).Visible = False
      lblDpat.Visible = False
   Case 1
      Caption = Majus(ID_Véhicule) & " " & ID_Articulé
      fraRemorque1.Visible = True
      fraRemorque2.Visible = False
      If contextBibliVeh Then txtDpat(0).Enabled = True
      txtDpat(1).Visible = False
  
   Case 2
      Caption = Majus(ID_Véhicule) & " " & ID_BiArticulé
      fraRemorque1.Visible = True
      fraRemorque2.Visible = True
      If contextBibliVeh Then
        txtDpat(0).Enabled = True
        txtDpat(1).Enabled = True
      End If
  End Select

  Illustration(0).Visible = True
  idxillustr = 0
   
  If cmdCancel.Caption = ID_Fermer Then
    ' pas de bouton OK
    cmdOK.Visible = False
    cmdCancel.Default = True
  End If
   
   If Not lblBraqReduit.Visible Then BraqReduitSignalé = False
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If contextBibliVeh Then
    If UnloadMode > vbFormCode Then   ' sinon l'appel vient de frmBibVehicule en vue d'une réouverture
        frmBibVéhicule.cmdTerminer = True    ' pour assurer la prise en compte des éventuelles modifs
    End If
  End If
  
End Sub

Private Sub cmdHelp_Click()

    SendKeys "{F1}", True
    
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim Cancel As Boolean

  If ActiveControl.name <> "cmdOK" Then
    SendKeys "{TAB}"
  
  Else

    For i = 0 To nbRemorque
      If Trim(txtPAFar(i)) = "" Then
        MsgBox ID_PAFar & " " & Idm_Obligatoire, vbYes + vbExclamation
        txtPAFar(i).SetFocus
        Cancel = True
        Exit For
      End If
    Next
  
    If Not Cancel Then
      If Len(Trim(txtaVehMax.Text)) = 0 Then
        MsgBox ID_VitBraq & " " & Idm_Obligatoire, vbYes + vbExclamation
        txtaVehMax.SetFocus
      ElseIf lblRayVehic(0).Caption = "" Then
        MsgBox Idm_Incomplet, vbYes + vbExclamation
      ElseIf Not calculOk Then
        MsgBox Idm_Incomplet, vbYes + vbExclamation
      Else
        
        If VersionDemo Then
          MsgBox ID_ExemplDemo & vbCrLf & IDm_PasDeMemo, vbExclamation
          CacheMoi
        Else
          CacheMoi
          EcrireVéhicule
        End If
        Modif = False
        
      End If
    End If
    
  End If

End Sub

'*****************************************************************************************************************************************
'             Saisie des caractéristiques géométriques
'*****************************************************************************************************************************************
Private Sub TraiterModification(ByVal uneTextBox As TextBox, ByVal ExValeur As String)
  
  If recup(uneTextBox.Text) <> recup(ExValeur) Then
    Modif = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print Cancel
End Sub

Private Sub txtNom_Validate(Cancel As Boolean)
  txtNom.Text = Trim(txtNom.Text)
  If Len(txtNom.Text) = 0 Then
    Cancel = True
  ElseIf ExistVeh(txtNom.Text, ColVéhicules) And txtNom.Text <> mVéhicule.Nom Then
    MsgBox "Ce véhicule existe déjà dans la bibliothèque"
    Cancel = True
  End If
End Sub

Private Sub txtPAFav_GotFocus(Index As Integer)
    
  VerifIllustr (Index)
  sauvPAFav(Index) = txtPAFav(Index)
  iniTextBox txtPAFav(Index)
  
End Sub

Private Sub txtPAFav_Validate(Index As Integer, Cancel As Boolean)
' Si Index=0 (Porte-à-faux Avant du tracteur) : valeur strictement positive
  If controleNumeric(txtPAFav(Index), StrictPositif:=(Index = 0)) Then
    txtPAFav(Index) = sauvPAFav(Index)
    Cancel = True
  ElseIf ctrlGeometrie(Index, txtPAFav(Index), sauvPAFav(Index)) Then
    txtPAFav(Index) = sauvPAFav(Index)
    Cancel = True
    If calculDéjaFait Then calcul (0)
  Else
    txtPAFav(Index) = affiche(txtPAFav(Index), True)
    TraiterModification txtPAFav(Index), sauvPAFav(Index)
  End If
End Sub

Private Sub txtE_GotFocus(Index As Integer)
  
  VerifIllustr (Index)
  sauvE(Index) = txtE(Index)
  iniTextBox txtE(Index)

End Sub

Private Sub txtE_Validate(Index As Integer, Cancel As Boolean)
  If controleNumeric(txtE(Index), StrictPositif:=True) Then
    txtE(Index) = sauvE(Index)
    Cancel = True
  ElseIf ctrlDpat(Index) Then
      txtE(Index) = sauvE(Index)
      Cancel = True
  ElseIf ctrlGeometrie(Index, txtE(Index), sauvE(Index)) Then
      txtE(Index) = sauvE(Index)
      Cancel = True
      If calculDéjaFait Then calcul (0)
  Else
    txtE(Index) = affiche(txtE(Index))
    TraiterModification txtE(Index), sauvE(Index)
  End If
End Sub

Private Sub txtPAFar_GotFocus(Index As Integer)
    
    VerifIllustr (Index)
    sauvPAFar(Index) = txtPAFar(Index)
    iniTextBox txtPAFar(Index)
    
End Sub

Private Sub txtPAFar_Validate(Index As Integer, Cancel As Boolean)
  If controleNumeric(txtPAFar(Index), StrictPositif:=True) Then
    txtPAFar(Index) = sauvPAFar(Index)
    Cancel = True
  Else
    txtPAFar(Index) = affiche(txtPAFar(Index))
    TraiterModification txtPAFar(Index), sauvPAFar(Index)
  End If
End Sub

Private Sub txtL_GotFocus(Index As Integer)
   VerifIllustr (Index)
  sauvL(Index) = txtL(Index)
  iniTextBox txtL(Index)
End Sub

Private Sub txtL_Validate(Index As Integer, Cancel As Boolean)
  If controleNumeric(txtL(Index), StrictPositif:=True) Then
    txtL(Index) = sauvL(Index)
    Cancel = True
  ElseIf ctrlGeometrie(Index, txtL(Index), sauvL(Index)) Then
    txtL(Index) = sauvL(Index)
    Cancel = True
    If calculDéjaFait Then calcul (0)
  Else
    txtL(Index) = affiche(txtL(Index))
    TraiterModification txtL(Index), sauvL(Index)
  End If
End Sub

Private Sub txtLEav_GotFocus()
  
  VerifIllustr 0
  sauvLEav = txtLEav
  iniTextBox txtLEav

End Sub

Private Sub txtLEav_Validate(Cancel As Boolean)
  If controleNumeric(txtLEav, StrictPositif:=True) Then
    txtLEav = sauvLEav
    Cancel = True
  ElseIf ctrlGeometrie(0, txtLEav, sauvLEav) Then
    txtLEav = sauvLEav
    Cancel = True
    If calculDéjaFait Then calcul (0)
  Else
    txtLEav = affiche(txtLEav)
    TraiterModification txtLEav, sauvLEav
  End If
End Sub

Private Sub txtLEar_GotFocus(Index As Integer)
  VerifIllustr (Index)
  sauvLEar(Index) = txtLEar(Index)
  iniTextBox txtLEar(Index)
End Sub

Private Sub txtLEar_Validate(Index As Integer, Cancel As Boolean)
   
   Select Case Index
   Case TRACTEUR
    If idxillustr <> 0 Then
      Illustration(idxillustr).Visible = False
      Illustration(0).Visible = True
      idxillustr = 0
    End If
  Case Else
    If idxillustr <> 2 Then
        Illustration(idxillustr).Visible = False
        Illustration(2).Visible = True
        idxillustr = 2
    End If
  End Select
  
  If controleNumeric(txtLEar(Index), StrictPositif:=True) Then
    txtLEar(Index) = sauvLEar(Index)
    Cancel = True
  ElseIf ctrlGeometrie(Index, txtLEar(Index), sauvLEar(Index)) Then
    txtLEar(Index) = sauvLEar(Index)
    Cancel = True
    If calculDéjaFait Then calcul (0)
  Else
    txtLEar(Index) = affiche(txtLEar(Index))
    TraiterModification txtLEar(Index), sauvLEar(Index)
  End If

End Sub

Private Sub txtDpat_GotFocus(Index As Integer)
  
  AffIllustr 1
  sauvDpat(Index) = txtDpat(Index)
  iniTextBox txtDpat(Index)
  
End Sub

Private Sub txtDpat_Validate(Index As Integer, Cancel As Boolean)
  If controleNumeric(txtDpat(Index), StrictPositif:=False) Then
    txtDpat(Index) = sauvDpat(Index)
    Cancel = True
  ElseIf ctrlDpat(Index) Then
    txtDpat(Index) = sauvDpat(Index)
    Cancel = True
  ElseIf ctrlGeometrie(Index + 1, txtDpat(Index), sauvDpat(Index), True) Then
    txtDpat(Index) = sauvDpat(Index)
    Cancel = True
    If calculDéjaFait Then calcul (0)
  Else
    txtDpat(Index) = affiche(txtDpat(Index), True)
    TraiterModification txtDpat(Index), sauvDpat(Index)
  End If
End Sub

Private Sub cmdImprim_Click()

Dim memoX As Single, memoY As Single, lgSpace As Single
Dim memoxdeb As Single, memoydeb As Single
Dim ColTrac%

         
'Me.Hide
      
' Copie d'écran
'PrintForm
'Me.Show
'Exit Sub

Me.MousePointer = vbHourglass
cmdImprim.Enabled = False

Printer.ScaleMode = 7
Printer.ScaleTop = -2
Printer.CurrentX = 5
Printer.CurrentY = -1.2
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print "GIRATION " & GirationVersion;
Printer.Print Spc(3); ID_MaitriseOuvrage
Printer.FontSize = 10
Printer.CurrentX = 5
Printer.CurrentY = -0.7
Printer.FontSize = 10
    Printer.FillStyle = 2
    Printer.Print ""
    Printer.FillStyle = vbFSTransparent
Printer.Print ""
Printer.ScaleLeft = -5
'memoxdeb = Printer.CurrentX
memoydeb = Printer.CurrentY

  '=== Géométrie ==
    ColTrac = Len(ElimEnTete(lblLEar + "  ")) + 2    ' Le +2 est pour tenir compte du Tab(2) en tête
    Printer.Print ""
    If nbRemorque = 0 Then
      Printer.Print Tab(2); Majus(ID_Véhicule) & " : " & lblNom; Tab(ColTrac - 2); Spc(1);
      memoX = Printer.CurrentX + Printer.TextWidth(Majus(ID_Tracteur, Tout:=True))
    Else
      Printer.Print Tab(2); Majus(ID_Véhicule) & " : " & lblNom; Tab(ColTrac - 2); fraTracteur; Tab(ColTrac + 8); fraRemorque1; Spc(1);
      If nbRemorque = 2 Then Printer.Print Tab(ColTrac + 20); fraRemorque2; Spc(1);
      memoX = Printer.CurrentX
    End If
    Printer.Print "": Printer.Print
    Printer.Print Tab(2); ElimEnTete(lblPAFav); Tab(ColTrac); txtPAFav(0); Tab(ColTrac + 10); txtPAFav(1); Tab(ColTrac + 22); txtPAFav(2)
    Printer.Print Tab(2); ElimEnTete(lblE); Tab(ColTrac); txtE(0); Tab(ColTrac + 10); txtE(1); Tab(ColTrac + 22); txtE(2)
    Printer.Print Tab(2); ElimEnTete(lblPAFar); Tab(ColTrac); txtPAFar(0); Tab(ColTrac + 10); txtPAFar(1); Tab(ColTrac + 22); txtPAFar(2)
    Printer.Print Tab(2); ElimEnTete(lblL); Tab(ColTrac); txtL(0); Tab(ColTrac + 10); txtL(1); Tab(ColTrac + 22); txtL(2)
    Printer.Print Tab(2); ElimEnTete(lblLEav); Tab(ColTrac); txtLEav
    Printer.Print Tab(2); ElimEnTete(lblLEar); Tab(ColTrac); txtLEar(0); Tab(ColTrac + 10); txtLEar(1); Tab(ColTrac + 22); txtLEar(2);
    If nbRemorque <> 0 Then
      Printer.Print ""
      Printer.Print Tab(2); ElimEnTete(lblDpat); Tab(ColTrac); txtDpat(0); Tab(ColTrac + 10); txtDpat(1);
    End If
    
    Printer.Print ""
    memoY = Printer.CurrentY
'    Printer.Line (memoxdeb, memoydeb)-(memoX, memoY), , B
    ImprimeCadre memoxdeb, memoydeb, memoX, memoY
    Printer.Print ""

  '=== Tracteur ==
    memoxdeb = Printer.CurrentX
    memoydeb = Printer.CurrentY
    Printer.Print ""
    Printer.Print Tab(2); fraCarTrac
    Printer.Print ""
    Printer.Print Tab(2); lblAVehMax; lblUnite; Spc(2); txtaVehMax;
'    Printer.CurrentX = Printer.CurrentX + 1.5 - Printer.TextWidth(txtaVehMax)
'    Printer.Print txtaVehMax; Spc(1);
'    ColTrac = Len(lblAVehMax & "  " & lblUnite + "  ") + 2    ' Le +2 est pour tenir compte du Tab(2) en tête
'    Printer.Print Tab(2); lblAVehMax; lblUnite; Tab(ColTrac); txtAVehMax;
    memoX = Printer.CurrentX
    Printer.Print Tab(2); cboRayon;
    Printer.CurrentX = memoX - Printer.TextWidth(txtRTracteur)
    Printer.Print txtRTracteur; Spc(1);
    memoX = Printer.CurrentX
    Printer.Print ""
'    Printer.Line (memoxdeb, memoydeb)-(memoX, Printer.CurrentY), , B
    ImprimeCadre memoxdeb, memoydeb, memoX, Printer.CurrentY
    
  '=== Véhicule ==
    Printer.Print ""
    memoxdeb = Printer.CurrentX
    memoydeb = Printer.CurrentY
    Printer.Print ""
    Printer.Print Tab(2); fraCarVeh
    Printer.Print ""
    If lblBraqReduit.Visible Then Printer.Print Tab(2); lblBraqReduit: Printer.Print ""
    Dim L1 As Single, L2 As Single, l3 As Single
    L1 = Max(Printer.TextWidth(lblRayInter), Printer.TextWidth(lblRayExter))
    L2 = Printer.TextWidth(lblEntreMurs) - 0.3
    l3 = Printer.TextWidth(lblEntreTrottoirs) - 0.3
    Printer.CurrentX = L1
    Printer.Print Spc(4); lblEntreMurs;
    L2 = Printer.CurrentX
    Printer.Print Spc(2); lblEntreTrottoirs;
    l3 = Printer.CurrentX
    Printer.Print Tab(2); lblRayInter;
    Printer.CurrentX = L2 - Printer.TextWidth(lblRayVehic(1))
    Printer.Print lblRayVehic(1);
    Printer.CurrentX = l3 - Printer.TextWidth(lblRayVehic(0))
    Printer.Print lblRayVehic(0); Spc(1);
    memoX = Printer.CurrentX
    Printer.Print Tab(2); lblRayExter;
    Printer.CurrentX = L2 - Printer.TextWidth(lblRayVehic(3))
    Printer.Print lblRayVehic(3);
    Printer.CurrentX = l3 - Printer.TextWidth(lblRayVehic(2))
    Printer.Print lblRayVehic(2); Spc(1)
'    Printer.Print Tab(22); lblEntreMurs; Tab(35); lblEntreTrottoirs;
'    memoX = Printer.CurrentX
'    Printer.Print Tab(2); lblRayInter; Tab(24); lblRayVehic(1); Tab(38); lblRayVehic(0)
'    Printer.Print Tab(2); lblRayExter; Tab(24); lblRayVehic(3); Tab(38); lblRayVehic(2)
'    Printer.Line (memoxdeb, memoydeb)-(memoX, Printer.CurrentY), , B
    ImprimeCadre memoxdeb, memoydeb, memoX, Printer.CurrentY

    Printer.EndDoc
    
  '  Me.Visible = True
    cmdImprim.Enabled = True
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ImprimeCadre(ByVal X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
    Printer.Line (X1, Y1)-(X2, Y2), , B
End Sub

Private Function ElimEnTete(label As String) As String

  ElimEnTete = Mid(label, InStr(label, ":") + 2)
  
End Function

Private Sub cboRayon_Click()     ' Réf: §2.a
  If txtRTracteur <> "" Then
  'Le rayon a déjà été saisi(ou affiché) au moins une fois : les autres ont été calculés
    txtRTracteur = affiche(R(cboRayon.ListIndex + 1, 0))
  End If
  lblMsgRExtEntreMurs.Visible = (cboRayon.ListIndex = 3) And Not txtRTracteur.Locked
End Sub

'*****************************************************************************************************************************************
'             Saisie du rayon de giration du tracteur et de la vitesse de braquage
'*****************************************************************************************************************************************

Private Sub txtRTracteur_GotFocus()
  If contextBibliVeh And chkVerrou = vbChecked Then SendKeys "{TAB}": Exit Sub
  AffIllustr 3
  sauvRVehMin = txtRTracteur
  iniTextBox txtRTracteur
  
End Sub

Private Sub txtRTracteur_Validate(Cancel As Boolean)

  If controleNumeric(txtRTracteur, StrictPositif:=True) Then
    txtRTracteur = sauvRVehMin
    Cancel = True
  ElseIf txtRTracteur <> "" And txtRTracteur <> sauvRVehMin Then
    txtRTracteur = affiche(txtRTracteur)
    BraqReduitSignalé = False
    calcul 0, True
    If Len(txtRTracteur) = 0 Then
      Cancel = True
    Else
      TraiterModification txtRTracteur, sauvRVehMin
    End If
  End If

End Sub

Private Sub txtAVehMax_GotFocus()
  If Not contextBibliVeh Or chkVerrou = vbChecked Then SendKeys "{TAB}": Exit Sub
  sauvaVehMax = txtaVehMax
  iniTextBox txtaVehMax
End Sub

Private Sub txtaVehMax_Validate(Cancel As Boolean)
Dim flagmax As Double, Msg As String

  If OptGen.unite = DEGRE Then
      flagmax = 15
  '    Msg = "15 degrés/seconde)"
  Else
    flagmax = 17
  '  Msg = "17 grades/seconde)"
  End If
  Msg = CStr(flagmax) & " " & libUnite(OptGen.unite) & "/" & ID_seconde

  If controleNumeric(txtaVehMax, StrictPositif:=True) Then
    txtaVehMax = sauvaVehMax
    Cancel = True
  ElseIf txtaVehMax = "" Then
  Else
    txtaVehMax = affiche(txtaVehMax)
    If CDbl(txtaVehMax) < 1 Then
      MsgBox "Minimum : 1 " & libUnite(OptGen.unite) & "/" & ID_seconde, vbYes + vbExclamation
      txtaVehMax = sauvaVehMax
      Cancel = True
    ElseIf CDbl(txtaVehMax) > OptGen.VitMax Then
      MsgBox "Maximum : " & CStr(OptGen.VitMax) & " " & libUnite(OptGen.unite) & "/" & ID_seconde, vbYes + vbExclamation
      txtaVehMax = sauvaVehMax
      Cancel = True
    ElseIf CDbl(txtaVehMax) > flagmax Then
      MsgBox IDm_ValeurGrande & " (> " & Msg, vbYes + vbExclamation
    End If
    If Not Cancel Then
      TraiterModification txtaVehMax, sauvaVehMax
    End If
  End If
  

End Sub

'*****************************************************************************************************************************************
'             Controles
'*****************************************************************************************************************************************
Private Function ctrl_alpha_max() As Boolean
Dim alphaVehMax As Double
                                                                                    ' Réf.: 1 et 3.a §4
  alphaVehMax = Atn(txtE(0) / (R(1, 0) + txtLEar(0) / 2))
  If alphaVehMax < pi / 20 Then
    MsgBox IDm_BraqTropFaible, vbYes + vbExclamation
    ctrl_alpha_max = True
  ElseIf alphaVehMax > pi / 3 Then
    MsgBox IDm_BraqTropFort, vbYes + vbExclamation
    ctrl_alpha_max = True
  Else
    ctrl_alpha_max = False
  End If
  
  End Function

Private Function ctrlDpat(Index As Integer) As Boolean
                                                                  ' Réf.: 3.b §2
  If Index = REMORQUE2 Or (Index = TRACTEUR And nbRemorque = 0) Then
    ctrlDpat = False
  ElseIf txtDpat(Index) = "" Or txtE(Index) = "" Then
    ctrlDpat = False
  ElseIf txtDpat(Index) <= -txtE(Index) / 2 Then
    MsgBox IDm_AttacheEloignée
    ctrlDpat = True
  Else
    ctrlDpat = False
  End If
End Function

Private Sub calcul(Index As Integer, Optional flagTracteur As Variant)
Dim i As Integer

  If Index = TRACTEUR Then       ' Tracteur
    lblBraqReduit.Visible = False
    calcule(TRACTEUR) = calcElem(Index)
    calcule(REMORQUE1) = False
    calcule(REMORQUE2) = False
  End If
      
  If nbRemorque > 0 And autorCalcul(REMORQUE1) And calcule(TRACTEUR) And Not calcule(REMORQUE1) Then
    calcule(REMORQUE1) = calcElem(REMORQUE1)
    calcule(REMORQUE2) = False
  End If

  If nbRemorque = 2 And autorCalcul(REMORQUE2) And calcule(REMORQUE1) And Not calcule(REMORQUE2) Then
    calcule(REMORQUE2) = calcElem(REMORQUE2)
  End If
        
  Select Case nbRemorque
    Case 0
        calculOk = calcule(TRACTEUR)
      If calculOk Then
        For i = 0 To 3
          lblRayVehic(i) = affiche(R(i + 1, 0))
        Next
      End If
    Case 1
      calculOk = calcule(TRACTEUR) And calcule(REMORQUE1)
      If calculOk Then
        For i = 0 To 1
          lblRayVehic(i) = affiche(Min(R(i + 1, 0), R(i + 1, 1)), True)
          If CDbl(lblRayVehic(i) < 0) Then lblRayVehic(i) = ID_SurPlace
        Next
        For i = 2 To 3
          lblRayVehic(i) = affiche(Max(R(i + 1, 0), R(i + 1, 1)), True)
          If CDbl(lblRayVehic(i) < 0) Then lblRayVehic(i) = ID_SurPlace
        Next
      End If
    Case 2
      calculOk = calcule(TRACTEUR) And calcule(REMORQUE1) And calcule(REMORQUE2)
      If calculOk Then
        For i = 0 To 1
          lblRayVehic(i) = affiche(Min(R(i + 1, 0), R(i + 1, 1), R(i + 1, 2)), True)
          If CDbl(lblRayVehic(i) < 0) Then lblRayVehic(i) = ID_SurPlace
        Next
        For i = 2 To 3
          lblRayVehic(i) = affiche(Max(R(i + 1, 0), R(i + 1, 1), R(i + 1, 2)), True)
          If CDbl(lblRayVehic(i) < 0) Then lblRayVehic(i) = ID_SurPlace
        Next
      End If
  End Select
  
  If Not calculOk Then
    For i = 0 To 3
      lblRayVehic(i) = ""
    Next
    lblBraqReduit.Visible = False
    
    If Index = 0 And Not calcule(TRACTEUR) And Not IsMissing(flagTracteur) Then
      txtRTracteur = ""
    End If
  End If
  
End Sub

'******************************************************************************
' Calcul des rayons d'un élément de carrosserie
' Index = 0 (Tracteur), 1 (Remorque1), 2 (Remorque2)
'******************************************************************************
Private Function calcElem(Index As Integer) As Boolean      ' Réf.: § 2 - Caractéristiques de giration
Dim i%, w As Double

' calcul de Rt,R1 et R2
  If Index = TRACTEUR Then     ' Tracteur  --> R(0,Index) correspond à Rt
    Select Case cboRayon.ListIndex
    Case 0
      R(0, Index) = txtRTracteur + txtLEar(Index) / 2
    Case 1
      R(0, Index) = txtRTracteur + txtL(Index) / 2
    Case 2
      On Error GoTo GestErr
      R(0, Index) = Sqr(Carre(txtRTracteur) - Carre(txtE(Index))) - txtLEav / 2
      On Error GoTo 0
    Case 3
      On Error GoTo GestErr
      R(0, Index) = Sqr(Carre(txtRTracteur) - Carre(txtE(Index) + CDbl(txtPAFav(Index)))) - txtL(Index) / 2
      On Error GoTo 0
    End Select
  
  Else     ' Remorque --> R(0,Index) correspond à R1 ou R2
    w = Carre(txtE(Index)) - Carre(txtDpat(Index - 1))
    If w <= Carre(R(0, Index - 1)) Then                  ' Réf: §4 - Calcul des caractéristiques du véhicule
      R(0, Index) = Sqr(Carre(R(0, Index - 1)) - w)
    Else
      lblBraqReduit.Visible = True
      If Not BraqReduitSignalé Then MsgBox IDm_RemorqueRayon: BraqReduitSignalé = True
      If Index = REMORQUE1 Then
        R(0, Index) = 0
        R(0, 0) = Sqr(w) + 0.001
        calcElem = calcElem(REMORQUE1)
      Else  ' Remorque2
        R(0, Index) = 0
        w = w + Carre(txtE(Index - 1)) - Carre(txtDpat(Index - 2))
        R(0, 0) = Sqr(w) + 0.01
        calcElem (REMORQUE1)
        calcElem = calcElem(REMORQUE2)
      End If
      Exit Function
    End If
  End If
                      
          ' Rayon intérieur entre trottoirs
  R(1, Index) = R(0, Index) - txtLEar(Index) / 2
  
          ' Rayon intérieur entre murs
  R(2, Index) = R(0, Index) - txtL(Index) / 2
          
          ' Rayon extérieur entre trottoirs
  If Index = TRACTEUR Then   ' Tracteur
    R(3, Index) = Sqr(Carre(R(0, Index) + txtLEav / 2) + Carre(txtE(Index)))
  Else                      ' Remorque
    R(3, Index) = R(0, Index) + txtLEar(Index) / 2
  End If
          
          ' Rayon extérieur entre murs
  R(4, Index) = Sqr(Carre(R(0, Index) + txtL(Index) / 2) + Carre(txtE(Index) + CDbl(txtPAFav(Index))))
  
  If Index = TRACTEUR Then
    For i = 1 To 4        ' Réf. §3a-Méthode de calcul
      If R(i, Index) < 0 Then
        MsgBox cboRayon.Text & " : " & IDm_TropFaible, vbYes + vbExclamation
        calcElem = False
        Exit Function
      End If
    Next
  End If
  
  If Index = 0 Then
    calcElem = Not ctrl_alpha_max
  Else
    calcElem = True
  End If
  
Exit Function

GestErr:
If Err = 5 Then       ' Plantage sur le calcul de racine d'un nb négatif (cborayon.ListIndex=2 ou 3)
  MsgBox cboRayon.Text & " : " & IDm_TropFaible, vbYes + vbExclamation
  txtRTracteur = sauvRVehMin: txtRTracteur.SetFocus
End If

End Function

Private Function ctrlGeometrie(Index As Integer, Valeur As String, ByVal ExValeur As String, Optional rayonConcerne As Variant) As Boolean
' Fonction appelée suite à une modification de la géométrie
' RayonConcerné est généralement vrai, sauf pour txtDpat

Dim i As Integer

  If Valeur = ExValeur Then Exit Function
  
  calculDéjaFait = calculOk
  calculOk = False
  calcule(Index) = False
  
  For i = 0 To 3                ' Réf.: 3.b §1
    lblRayVehic(i) = ""
  Next
  lblBraqReduit.Visible = False
  
  If Index = TRACTEUR Then
    If IsMissing(rayonConcerne) Then txtRTracteur = ""          ' Réf.: 3.b §1
    If autorCalcul(Index) Then
       txtRTracteur.Enabled = True   ' Réf.: 3.a §1
    Else
      txtRTracteur.Enabled = False
    End If
  
  ElseIf txtRTracteur <> "" And autorCalcul(Index) Then     ' Réf.: 3.b §3
    calcul Index
    If Not calcule(Index) Then
      ctrlGeometrie = True
    End If
  End If

End Function

Private Function autorCalcul(Index As Integer)
    
  autorCalcul = txtPAFav(Index) <> "" And txtE(Index) <> "" And txtL(Index) <> "" And txtLEar(Index) <> ""
  If Index = TRACTEUR Then
    autorCalcul = autorCalcul And txtLEav <> ""
  Else
    autorCalcul = autorCalcul And txtDpat(Index - 1) <> "" And calcule(Index - 1)
  End If
  
End Function

Private Sub EcrireVéhicule()

Dim numRemorque As Integer
  
On Error GoTo GestErr

  With VehicTab
    .Nom = txtNom.Text
    .Protégé = (chkVerrou = vbChecked)
    .nbRemorque = nbRemorque
    For numRemorque = TRACTEUR To REMORQUE2
      .PAFav(numRemorque) = recup(txtPAFav(numRemorque))
      .e(numRemorque) = recup(txtE(numRemorque))
      .PAFar(numRemorque) = recup(txtPAFar(numRemorque))
      .l(numRemorque) = recup(txtL(numRemorque))
      .LEar(numRemorque) = recup(txtLEar(numRemorque))
      If numRemorque < REMORQUE2 Then
        .Dpat(numRemorque) = recup(txtDpat(numRemorque))
      End If
    Next
    .LEav = recup(txtLEav)
    .aVehMax = angConv(recup(txtaVehMax), radian)
    .RVehMin = R(1, 0)
  End With

  If Creation Then
    'Dim NewVehic As VEHICULE
    'Set NewVehic = InserVéhicule(VehicTab)
    frmBibVéhicule.AjouterVéhicule InserVéhicule(VehicTab)
    Creation = False
    
  Else
  ' Mise  à jour de l'existant
    With ColVéhicules(UCase(lblNom))
      .Protégé = VehicTab.Protégé
        For numRemorque = 0 To nbRemorque
          .PAFav(numRemorque) = VehicTab.PAFav(numRemorque)
          .e(numRemorque) = VehicTab.e(numRemorque)
          .PAFar(numRemorque) = VehicTab.PAFar(numRemorque)
          .l(numRemorque) = VehicTab.l(numRemorque)
          .LEar(numRemorque) = VehicTab.LEar(numRemorque)
          If numRemorque > 0 Then
            .Dpat(numRemorque - 1) = VehicTab.Dpat(numRemorque - 1)
          End If
        Next
        .LEav = VehicTab.LEav
      .alphaVehMax = VehicTab.alphaVehMax
      .aVehMax = VehicTab.aVehMax
      .RVehMin = VehicTab.RVehMin
    End With
    
    If UCase(lblNom.Caption) <> UCase(txtNom.Text) Then
    'changement de clé
      frmBibVéhicule.RenommerVéhicule InserVéhicule(VehicTab)
    End If
  End If
  
  With frmBibVéhicule
    If chkVerrou = vbChecked Then
      .cmdAction(1).Caption = ID_Visualiser & "..."
      .cmdAction(2).Enabled = False
    Else
      .cmdAction(1).Caption = ID_Modifier & "..."
      .cmdAction(2).Enabled = True
    End If
    .Modif = True
  End With
    
Exit Sub
GestErr:
  If Err <> ErrVéhicule Then
    ErreurNonFatale "EcrireVéhicule"
  End If

  
End Sub

Private Function recup(v As String) As Double
  If v = "" Then
    recup = 0#
  Else
    recup = CDbl(v)
  End If
End Function

Private Function affiche(ByVal v As Variant, Optional ByVal flag As Boolean) As Variant
  If v = "" Then
    affiche = ""
  
  ElseIf v = 0# And Not flag Then
    affiche = ""
  
  Else
    'affiche = Format(v, "####0.000")
    ' v3.3
    affiche = Format(v, "####0.00")
  End If

End Function

Private Sub AfficherVéhicule(Vehic As VEHICULE)

Dim numRemorque As Integer
  
With Vehic
  
  For numRemorque = 0 To nbRemorque
    txtPAFav(numRemorque) = affiche(.PAFav(numRemorque), flag:=nbRemorque > 0)
    txtE(numRemorque) = affiche(.e(numRemorque))
    txtPAFar(numRemorque) = affiche(.PAFar(numRemorque))
    txtL(numRemorque) = affiche(.l(numRemorque))
    txtLEar(numRemorque) = affiche(.LEar(numRemorque))
    If numRemorque > 0 Then
      txtDpat(numRemorque - 1) = affiche(.Dpat(numRemorque - 1), True)
    End If
  Next
  
  txtLEav = affiche(.LEav)
  txtaVehMax = affiche(angConv(.aVehMax, unite))
  txtRTracteur = affiche(.RVehMin)

End With
  
calcul TRACTEUR

End Sub

Private Sub chkVerrou_Click()

  If chkVerrou = vbChecked Then
    fraCarGeom.Enabled = False
    ' Simulation possible sur le rayon dans le contexte Trajectoire
    txtaVehMax.Locked = True: txtaVehMax.TabStop = False: txtaVehMax.CausesValidation = False
    If contextBibliVeh Then txtRTracteur.Locked = True: txtRTracteur.TabStop = False
    Me.txtNom.Locked = True
  Else
    fraCarGeom.Enabled = True
    txtaVehMax.Locked = False: txtaVehMax.TabStop = True
    If contextBibliVeh Then txtRTracteur.Locked = False: txtRTracteur.TabStop = True: txtRTracteur.CausesValidation = True
    Me.txtNom.Locked = False
    Me.txtNom.Visible = True
  End If

End Sub

Private Sub AffIllustr(numIllustr As Integer)

  If idxillustr <> numIllustr Then
    Illustration(idxillustr).Visible = False
    idxillustr = numIllustr
    Illustration(idxillustr).Visible = True
  End If

End Sub

Private Sub VerifIllustr(Index As Integer)

If Index = 0 Then
  AffIllustr 0
Else
  AffIllustr 2
End If

End Sub

