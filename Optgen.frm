VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options générales"
   ClientHeight    =   5280
   ClientLeft      =   960
   ClientTop       =   3135
   ClientWidth     =   7185
   Icon            =   "Optgen.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBibVeh 
      Caption         =   "Bibliothèque de véhicules"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   6855
      Begin VB.TextBox txtRepert 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmdParcourir 
         Caption         =   "Parcourir..."
         Height          =   405
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraRépertoires 
      Caption         =   "Dossiers par défaut"
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   6855
      Begin VB.TextBox untxtChemin 
         Height          =   285
         Left            =   360
         TabIndex        =   20
         Text            =   "textbox cachée"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox lstDossier 
         Height          =   450
         ItemData        =   "Optgen.frx":000C
         Left            =   240
         List            =   "Optgen.frx":000E
         TabIndex        =   10
         Top             =   360
         Width           =   6375
      End
      Begin VB.CommandButton cmdModifier 
         Caption         =   "Modifier..."
         Height          =   405
         Left            =   5520
         TabIndex        =   0
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7185
      TabIndex        =   17
      Top             =   4785
      Width           =   7185
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   320
         Left            =   3120
         TabIndex        =   7
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuler"
         CausesValidation=   0   'False
         Height          =   320
         Left            =   4440
         TabIndex        =   8
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Aide"
         Height          =   320
         Left            =   5760
         TabIndex        =   9
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraUnitéAngle 
      Caption         =   "Unité d'angle"
      Height          =   1332
      Left            =   5640
      TabIndex        =   16
      Top             =   3240
      Width           =   1332
      Begin VB.OptionButton optAngle 
         Caption         =   "Grade"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   852
      End
      Begin VB.OptionButton optAngle 
         Caption         =   "Degré"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   852
      End
   End
   Begin VB.Frame fraVitesses 
      Caption         =   "Vitesses"
      Height          =   1332
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   4692
      Begin VB.TextBox txtVitesse 
         Alignment       =   1  'Right Justify
         Height          =   288
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Top             =   840
         Width           =   492
      End
      Begin VB.TextBox txtVitesse 
         Alignment       =   1  'Right Justify
         Height          =   288
         Index           =   0
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   492
      End
      Begin VB.Label lblUnité 
         Caption         =   "Km/h"
         Height          =   252
         Left            =   3240
         TabIndex        =   15
         Top             =   840
         Width           =   492
      End
      Begin VB.Label angleSec 
         Height          =   252
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label lblVitesse 
         Caption         =   "Vitesse du véhicule par défaut"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   2172
      End
      Begin VB.Label lblVitBraq 
         Caption         =   "Limite supérieure admissible pour les vitesses de braquage"
         Height          =   492
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2292
      End
   End
   Begin MSComDlg.CommonDialog dlgBibVeh 
      Left            =   5040
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmOptGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmOptgen   -   OPTGEN.FRM
'   Feuille permettant la saisie des options générales du programme (config du poste de travail)
' Non accessible dès qu'une trajectoire est ouverte ou que l'on travaille sur la bibliothèque de véhicules

'**************************************************************************************
Option Explicit

Private tmpRepert(2) As String, sauvVit(1) As String
Public RepEnCours As Integer
Public untxtDossier As VB.TextBox

Private Sub cmdModifier_Click()
Dim sTitre As String
Dim bif_flag As Long

  RepEnCours = lstDossier.ListIndex
  
  'Appel de l'explorateur
  bif_flag = BIF_NEWDIALOGSTYLE
  'bif_flag = bif_flag Or BIF_RETURNFSANCESTORS
  untxtChemin.Text = Repert(RepEnCours)
  sTitre = Left(lstDossier, InStr(frmOptGen.lstDossier, vbTab) - 1)

  If Explorer.fso Is Nothing Then Set Explorer.fso = gtFso
  Explorer.Browse BIF_FLAGS:=bif_flag, sTitre:=sTitre, Feuille:=Me, Chemin:=Repert(RepEnCours)
  Repert(RepEnCours) = untxtChemin.Text
  
  'v3.3 : Instruction ci-dessous remplacée par l'explorateur ci-dessus + sympathique
'  frmExplorer.Show vbModal

End Sub

Private Sub cmdParcourir_Click()

Dim Cancel As Boolean
On Error GoTo ErrHandler

With dlgBibVeh
  .InitDir = txtRepert
  ' Bibliothèque de véhicules : flag FileMustExist et non seulement PathMustExist
'  .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNExplorer Or cdlOFNAllowMultiselect
  .flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNExplorer
  .FileName = gtFso.BuildPath(txtRepert, GIRATIONVEHCOURT)
  .ShowOpen
  If Not Cancel Then
    If StrComp(.FileTitle, GIRATIONVEHCOURT, vbTextCompare) <> 0 Then
      MsgBox IDm_NomBibvehInvalide
    Else
      Repert(2) = supprSlash(extraiRep(.FileName))
      If Not ExistFich(.FileName) Then MsgBox ID_Bibveh & " " & IDm_Absente
    End If
  End If
End With

Exit Sub

ErrHandler:   ' L'utilisateur a fait 'Annuler
  If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
  ElseIf Err.Number = cdlInvalidFileName And Right(txtRepert, 1) = "\" Then
    dlgBibVeh.FileName = txtRepert & GIRATIONVEHCOURT
    Resume
  Else
    ErrLocal
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  'v3.3 : Remplacement de l'explorateur
  'Unload frmExplorer
End Sub

Private Sub optAngle_Click(Index As Integer)
  
  If Index = 0 Then '  grade --> degré
    If txtVitesse(0) <> "" Then txtVitesse(0) = Format(txtVitesse(0) * 0.9, "##.0")
  Else                    ' degré --> grade
    If txtVitesse(0) <> "" Then txtVitesse(0) = Format(txtVitesse(0) / 0.9, "##.0")
  End If
  angleSec = libUnite(Index) & "/" & ID_seconde
  
End Sub

Private Sub cmdCancel_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()
  
'  DégraisserFonte Me
  Dim Feuille As Form
  Dim i As Integer
  
  'Affichage centré de la fenêtre
  Me.ScaleMode = 1
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  
  ' Aide contextuelle
  Me.HelpContextID = OPTGENERAL
  
  ' Désactivation des angles si une feuille est chargée
  For Each Feuille In Forms
    If Feuille.name = "frmTraject" Then
      optAngle(0).Enabled = False
      optAngle(1).Enabled = False
      Exit For
    End If
  Next
  
  With OptGen
    lstDossier.AddItem "Fonds de plan"
    lstDossier.AddItem "Trajectoires"
'    lstDossier.AddItem "Bibliothèque de véhicules"
    For i = 0 To 2
      Repert(i) = .Repert(i)
    Next
    optAngle(.unite) = True
    txtVitesse(0) = Format(.VitMax, "##.0")
    txtVitesse(1) = Format(.VitDef, "##.0")
  
    lstDossier.ListIndex = 0
    If ExistFich(.Repert(0), vbDirectory) Then
      If Not ExistFich(.Repert(1), vbDirectory) Then
        lstDossier.ListIndex = 1
      ElseIf Not ExistFich(.Repert(2), vbDirectory) Then
        txtRepert.TabIndex = 0 'lstDossier.ListIndex = 2
      End If
    End If
  End With
  
End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

Private Sub cmdOK_Click()
Dim Vehic As VEHICULE
Dim i As Integer

  For i = 0 To 2
    If Not ExistFich(tmpRepert(i), vbDirectory) Then
      If i = 2 Then
        txtRepert.SetFocus
      Else
        lstDossier.Selected(i) = True
      End If
      If Not CreateDossier(tmpRepert(i)) Then Exit Sub
    End If
  Next
  
 If VerifDroitBibVeh(tmpRepert(2)) Then Unload Me: Exit Sub

  With OptGen
    If .unite <> Numopt(optAngle) Then
      .unite = Numopt(optAngle)
    End If

    .Repert(0) = tmpRepert(0)   ' en principe superflu
    .Repert(1) = tmpRepert(1)   ' en principe superflu
    .VitMax = txtVitesse(0)
    .VitDef = txtVitesse(1)
    
    If StrComp(.Repert(2), tmpRepert(2), vbTextCompare) <> 0 Then
      .Repert(2) = tmpRepert(2)
      GirationVeh = gtFso.BuildPath(.Repert(2), GIRATIONVEHCOURT)
      If GirInitOk Then lireVeh
    End If

    SaveSetting Appname:=App.Title, SECTION:="Version", _
      Key:="V", Setting:=GirationVersion
  
    SaveSetting Appname:=App.Title, SECTION:="Parametres", _
      Key:="Unite", Setting:=Numopt(optAngle)
    SaveSetting Appname:=App.Title, SECTION:="Parametres", _
      Key:="FDP", Setting:=.Repert(0)
    SaveSetting Appname:=App.Title, SECTION:="Parametres", _
      Key:="Trajectoire", Setting:=.Repert(1)
    SaveSetting Appname:=App.Title, SECTION:="Parametres", _
      Key:="Véhicules", Setting:=.Repert(2)
    SaveSetting Appname:=App.Title, SECTION:="Parametres", _
      Key:="Vitesse max", Setting:=substPtDecimalRegional(txtVitesse(0), Regional:=False)
  '    Key:="Vitesse max", Setting:=txtVitesse(0)
    SaveSetting Appname:=App.Title, SECTION:="Parametres", _
      Key:="Vitesse defaut", Setting:=substPtDecimalRegional(txtVitesse(1), Regional:=False)
  '    Key:="Vitesse defaut", Setting:=txtVitesse(1)
    
    If GirInitOk Then
      gtRepertFDP = .Repert(0)
      MDIGiration.dlgTrajectoire.InitDir = .Repert(1)
    Else
      GirInitOk = True
    End If
    
  End With  ' Optgen
  
  Unload Me

End Sub

Private Sub txtRepert_GotFocus()

  'sauvRepert(Index) = txtRepert(Index)
   SendKeys "{HOME}+{END}"
End Sub

Private Sub txtRepert_Validate(Cancel As Boolean)
  
  If Trim(txtRepert) = "" Then
    MsgBox ID_Saisie & " " & Idm_Obligatoire
    Cancel = True
    txtRepert = tmpRepert(2)
    Exit Sub
  End If
    
  
  If StrComp(nomCourt(txtRepert, SansExtension:=False), GIRATIONVEHCOURT, vbTextCompare) = 0 Then txtRepert = extraiRep(txtRepert)
    
  If Not ExistFich(txtRepert, vbDirectory) Then
    If Not CreateDossier(txtRepert) Then
'      Cancel = True
      txtRepert = tmpRepert(2)
      Exit Sub
    End If
  End If
  
  Repert(2) = txtRepert
  
  'If Not ExistFich(txtRepert & "\" & GIRATIONVEHCOURT) Then
  '  MsgBox ID_Bibveh & IDm_Absente
  'End If
  
 
End Sub

Private Sub txtVitesse_Gotfocus(Index As Integer)
  sauvVit(Index) = txtVitesse(Index)
End Sub

' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la fonction KeyPress semble en mesure de réafficher la virgule

Private Sub txtVitesse_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True
'    If flagVirgule Then alertVirgule = True
'    If flagVirgule Then KeyCode = 188
'  End If

End Sub

Private Sub txtVitesse_KeyPress(Index As Integer, KeyAscii As Integer)
  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False
End Sub

Private Sub txtVitesse_Validate(Index As Integer, Cancel As Boolean)

Dim controle As Boolean, v As String
      
      v = txtVitesse(Index)
' Mis en commentaire le 23/02/2000 : on reste homogène avec les autres saisies (on tape soit le point décimal, soit le caractère du panneau de config)
' et de + çà pouvait buguer sur l'affectation finale (Fonction Format...)
'      If flagVirgule Then v = substVirgulePoint(v)

      If v = "" Then
        controle = False
      ElseIf Not IsNumeric(v) Then
        MsgBox IDm_Numerique
        controle = True
      ElseIf v = 0 Then
        MsgBox IDm_DifferentDeZero
        controle = True
      Else
        Select Case Index
        Case 0
          If v < 0 Then
            MsgBox IDm_SuperieurAZero
            controle = True
          End If
        Case 1
          If Abs(v) > 50 Then
            MsgBox IDm_MaxVitesse & " 50 " & ID_UniteVitesse
            controle = True
          ElseIf Abs(v) > 20 Then
            'controle = (MsgBox(ID_Vitesse & " > 20 " & ID_UniteVitesse & vbCrLf & vbCrLf & "      Confirmer ?", vbYesNo + vbDefaultButton2, IDm_Confirmation)) = vbNo
            controle = (MsgBox(ID_Vitesse & " > 20 " & ID_UniteVitesse & vbCrLf & vbCrLf, vbYesNo + vbQuestion + vbDefaultButton2, IDm_Confirmation)) = vbNo
          End If
        End Select
      End If
      
  If controle Then
    txtVitesse(Index) = sauvVit(Index)
    Cancel = True
  Else
    txtVitesse(Index) = Format(txtVitesse(Index), "##.0")
  End If
  
End Sub

Private Function CreateDossier(ByVal Dossier As String) As Boolean

  On Error GoTo GestErr
  
  If Dossier = "" Or ExistFich(Dossier, vbDirectory) Then
    MsgBox IDm_CreationDossierRefusee
  Else
    If MsgBox(IDm_CreerDossier & Dossier, vbOKCancel + vbQuestion + vbDefaultButton1) = vbOK Then
      MkDir Dossier
      CreateDossier = True
    End If
  End If
  
  Exit Function
  
GestErr:
  'Tentative de créer un dossier :
  ' Erreur 75 : "Erreur dans le chemin d'accès"
  '     - sans permission
  '     - existant (ou un nom de fichier existant)
  ' Erreur 76 : "Chemin d'accès introuvable"
  '    - sans que le dossier parent n'existe encore
  '    - avec un nom DOS incorrect (idem erreur 52 pour les fichiers)
  If Err = 76 Or Err = 75 Then
    If CreateDossier(gtFso.GetParentFolderName(Dossier)) Then Resume
  Else
    ErrGeneral "frmOptgen : CreateDossier"
  End If
  
End Function

Property Let Repert(ByVal Index As Integer, ByVal Value As String)
Dim Chaine As String
Dim chainecomplet As String
Dim pos As Integer

  tmpRepert(Index) = Value

  Select Case Index
  Case 0
    Chaine = "Fonds de plan" & vbTab & vbTab
    
  Case 1
    Chaine = "Trajectoires" & vbTab & vbTab
  Case 2
'    chaine = "Bibliothèque de véhicules" & vbTab
    txtRepert = Value
   If ExistFich(tmpRepert(Index), vbDirectory) And Not ExistFich(gtFso.BuildPath(tmpRepert(Index), GIRATIONVEHCOURT)) Then MsgBox ID_Bibveh & " " & IDm_Absente
   Exit Property
 End Select
  
  
  ' En principe, on affiche tout
  chainecomplet = Chaine & tmpRepert(Index)
  
  ' Rajout du répertoire racine
  Chaine = Chaine & Left(Value, 3)
  ' Positnt juste après le 1er anti-slash
  pos = 4
  
  ' Tronquage du chemin si trop long
  While (TextWidth(chainecomplet) - lstDossier.Width > -300) And pos <> 0
    pos = InStr(pos, Value, "\")
    If pos <> 0 Then
      pos = pos + 1
      chainecomplet = Chaine & "...\" & Mid(Value, pos)
    End If
  Wend
  
  lstDossier.List(Index) = chainecomplet

'  If Index = 2 And Not ExistFich(ajoutSlash(tmpRepert(Index)) & GIRATIONVEHCOURT) Then MsgBox ID_Bibveh & " " & IDm_Absente

End Property

Property Get Repert(ByVal Index As Integer) As String
  Repert = tmpRepert(Index)
End Property

