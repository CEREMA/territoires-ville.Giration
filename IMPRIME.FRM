VERSION 5.00
Begin VB.Form frmImprim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impression"
   ClientHeight    =   4170
   ClientLeft      =   2145
   ClientTop       =   2700
   ClientWidth     =   6150
   Icon            =   "IMPRIME.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboEpaisseur 
      Height          =   315
      ItemData        =   "IMPRIME.frx":000C
      Left            =   3720
      List            =   "IMPRIME.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   320
      Left            =   4680
      TabIndex        =   12
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CheckBox chkNomFichiers 
      Caption         =   "Nom des fichiers"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CheckBox chkTitre 
      Caption         =   "Titre de l'édition"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   1332
   End
   Begin VB.CommandButton cmdChoixImpr 
      Caption         =   "Imprimante ..."
      Height          =   320
      Left            =   4680
      TabIndex        =   11
      Top             =   2040
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      CausesValidation=   0   'False
      Height          =   320
      Left            =   4680
      TabIndex        =   9
      Top             =   1080
      Width           =   1332
   End
   Begin VB.TextBox txtTitre 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   4815
   End
   Begin VB.CommandButton cmdOptAff 
      Caption         =   "Options ..."
      Height          =   320
      Left            =   4680
      TabIndex        =   10
      Top             =   1560
      Width           =   1332
   End
   Begin VB.Frame fraEchelle 
      Caption         =   " Echelle "
      Height          =   1095
      Left            =   2400
      TabIndex        =   16
      Top             =   1440
      Width           =   1815
      Begin VB.TextBox txtEchelle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkAjuster 
         Caption         =   "Ajuster"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Label lblPrefixEchel 
         Caption         =   "1/"
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame fraNatImpr 
      Caption         =   " Imprimer "
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1815
      Begin VB.OptionButton optNatImpr 
         Caption         =   "Epure"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1332
      End
      Begin VB.OptionButton optNatImpr 
         Caption         =   "Caractéristiques"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1572
      End
   End
   Begin VB.Label lblEpaisseur 
      Caption         =   "Epaisseur de ligne"
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   1080
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblNomImpr 
      Caption         =   "Imprimante en cours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   735
      Index           =   1
      Left            =   1200
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblOrientation 
      Caption         =   "Orientation"
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmImprim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmImprim   -   IMPRIME.FRM
'   Saisir les paramètres d'impression et  la lancer

'**************************************************************************************
Option Explicit

Private basGauche As New PT, hautDroit As New PT
Private NomVehic As String
Private sauvEchelle As String, memoEchelAjustee As Long
' Pour faire le clipping via une fonction API, il faut être en pixel et donc faire une conversion
Dim facteur As Double, wPixel As Integer, hPixel As Integer

#If Win16 Then
    Private Declare Function ExcludeClipRect Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
#Else
   Private Declare Function ExcludeClipRect Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
#End If

Private Sub chkAjuster_Click()

  If chkAjuster = 1 Then
    txtEchelle = memoEchelAjustee
    ' Le double appel suivant permet de faire la mise à jour automatique du champ LibelTitre
    txtEchelle.SetFocus
    chkAjuster.SetFocus
    
    txtEchelle.Enabled = False
  Else
    
    txtEchelle.Enabled = True
  End If
  
End Sub

Private Sub cmdCancel_Click()
    
Unload Me
    
End Sub


Private Sub cmdChoixImpr_Click()
''Dim Cancel As Boolean
''
''  On Error GoTo ErrHandler
''
''  PrintDialog.flags = cdlPDPrintSetup   ' Pour afficher directement la fenêtre Configuration
''  PrintDialog.ShowPrinter

  Dim Cancel As Integer
  
  ShowPrinter Me, Cancel
  
  If Not Cancel Then
'    lblNomImpr = ID_ImprimanteEnCours & " " & Printer.DeviceName
    lblNomImpr = ID_ImprimanteEnCours & " " & TronqueChaine(Printer.DeviceName, 35)
    If Printer.Orientation = vbPRORLandscape Then
      Shape1(2).Visible = True
      Shape1(1).Visible = False
    Else
      Shape1(1).Visible = True
      Shape1(2).Visible = False
    End If
    Printer.ScaleMode = 7 ' cm
'    Printer.ScaleMode = 6 ' mm
   EchelAjustee
  End If
  
  
''  Exit Sub
''
''ErrHandler:     ' L'utilisateur a fait 'Annuler
''  Cancel = True
''  Resume Next

End Sub

Private Sub txtEchelle_KeyDown(KeyCode As Integer, Shift As Integer)
' on détecte la frappe du point décimal
' la fonction KeyPress annulera la frappe
  
  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True
'    alertVirgule = True
'  End If

End Sub

Private Sub txtEchelle_KeyPress(KeyAscii As Integer)
' Annulation de la frappe du point décimal
  If alertVirgule Then KeyAscii = 0: alertVirgule = False
  
End Sub

Private Sub txtEchelle_GotFocus()

  sauvEchelle = txtEchelle
  iniTextBox txtEchelle
  
End Sub

Private Sub txtEchelle_Validate(Cancel As Boolean)

  If controleNumeric(txtEchelle, StrictPositif:=True) Then
    txtEchelle = sauvEchelle
    Cancel = True
  Else
  ' Mise à jour du libellé du titre
    If txtEchelle = CStr(memoEchelAjustee) Or txtEchelle = "" Then
      txtTitre = NomVehic & _
                      "   " & Arrond(fCourante.v * 3.6, 2) & " " & ID_UniteVitesse
    Else
      txtTitre = NomVehic & _
                      "   " & Arrond(fCourante.v * 3.6, 2) & " " & ID_UniteVitesse & _
                      "   1/" & Arrond(CDbl(txtEchelle), 0) & "e"
    End If
  End If

End Sub

Private Sub Form_Load()

On Error GoTo GestErr

If FichierJournal Then Write #numFichLog, "Ouverture Impression"

'    DégraisserFonte Me

    'Affichage centré de la fenêtre
    Me.ScaleMode = 1
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
  ' Aide contextuelle
  Me.HelpContextID = IMPRIMER
    
  ' Recherche s'il y a une imprimante par défaut pout GIRATION
  InitializePrinter
  
    ' orientation en cours
  On Error GoTo GestErrOrientation
    Shape1(Printer.Orientation).Visible = True
  
    ' Option Epure par défaut
    optNatImpr(1) = True
    
    ' Noms de fichiers
    If fCourante.dlgTrajectoire.FileName = "" And fCourante.FDP = "" Then chkNomFichiers.Enabled = False

lblNomImpr = ID_ImprimanteEnCours & " " & Printer.DeviceName
cboEpaisseur = OptGen.Epaisseur
NomVehic = Trim(fCourante.Vehic.Nom)

If VersionDemo Then
  chkAjuster.Enabled = False
  chkTitre.Enabled = False
  txtTitre.Enabled = False
  txtTitre = ID_TitreImpressionVersionDemo
Else
  txtTitre = NomVehic & _
                      "   " & Arrond(fCourante.v * 3.6, 2) & " " & ID_UniteVitesse
End If
  
    basGauche.X = fCourante.pmin.X
    basGauche.Y = fCourante.pmin.Y
    hautDroit.X = fCourante.pmax.X
    hautDroit.Y = fCourante.pmax.Y
    
    Set ObjetDessin = Printer
    EchelAjustee
    
    Exit Sub
    
GestErrOrientation:
  Resume Next
  
GestErr:
  
  ErrGeneral
  
End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Unload frmAnnulImpr
  
  Printer.ScaleMode = 1   ' Retour auw Twips
  flagImpr = False
  
  ' Restitution des paramètres d'échelle de picTrajectoire
    With fCourante
      .Milieux.Remove .Milieux.Count
      .Echelles.Remove .Echelles.Count
'v3.3 : Mise en commentaire(02/05/07)
'ne sert à rien et en + par effet de bord peut générer l'erreur 5
'      .Form_Activate
'remplacé par la ligne suivante (04/06/07) : génère l'erreur 13 si un outil de mesure est sélectionné
      Set ObjetDessin = .picTrajectoire
    End With
    
    
End Sub

Private Sub optNatImpr_Click(Index As Integer)
' Nature de l'impression : Grille ou épure
Static memoEchelle As String

  If VersionDemo Then Exit Sub
  
  If Numopt(optNatImpr) = 1 Then
  ' Epure
    If Not VersionDemo Then chkAjuster.Enabled = True
    majTitre Index
    txtEchelle.Enabled = (chkAjuster = 0)
    cboEpaisseur.Enabled = True
  Else
  ' Grille
    chkAjuster.Enabled = False
    memoEchelle = txtEchelle
    txtEchelle = ""
    majTitre Index
    txtEchelle.Enabled = False
    txtEchelle = memoEchelle
    cboEpaisseur.Enabled = False
  End If
    
End Sub

Private Sub majTitre(Index As Integer)
' Mise à jour du titre suite à changement de choix d'impression (optNatImpr)

If Not txtEchelle.Visible Then Exit Sub    ' Feuille en cours de chargement

    txtEchelle.Enabled = True
    DoEvents
    ' Le double appel suivant permet de faire la mise à jour automatique du champ LibelTitre
'    Echelle.SetFocus
'    optNatImpr(Index).SetFocus
    DoEvents  ' Impératif

End Sub

Public Sub cmdOK_Click()
' déclenchement de l'impression de la grille ou du dessin
Dim Grille As Boolean
Dim w As Single, h As Single
Dim Titre1 As String, Titre2 As String
Dim Abandon As Boolean

''If ctrlRC(Me, "Echelle") Then Exit Sub
      
      ' Choix de l'impression ( grille ou dessin)
  Grille = (Numopt(optNatImpr) = 0)
  
       ' Définition des 2 lignes d'en-tête
  Entete Titre1, Titre2
  
  Me.Hide
  
  
  If Grille Then
    
    If FichierJournal Then Write #numFichLog, "Impression de la grille"

    With fCourante.Grille
    ' En-tête
                                        ' /fz"nn": taille en pouces  /fb0: non gras
      .PrintHeader = "/fz""12""/c" & Titre1 & "/n/n/fz""10""/c"
      .PrintHeader = .PrintHeader & Titre2
      .PrintHeader = .PrintHeader & "/n/n/n/fb1Origine : X=" & CStr(fCourante.txtOri(0)) & "    Y=" & CStr(fCourante.txtOri(1))
      If CDbl(fCourante.txtOri(2)) <> 0 Then .PrintHeader = .PrintHeader & "     " & ID_Direction & " = " & CStr(fCourante.txtOri(2)) & " " & libUnite(OptGen.unite)
      .PrintHeader = .PrintHeader & "/n"
      
    ' Marges
      .PrintMarginLeft = 1500
      .PrintMarginRight = 1500

  ' Impression de la grille
      .Action = 32
    End With
    
  Else  ' Dessin de l'Epure
    On Error GoTo GestErr
    Printer.Font.name = "Roman"
    
    ' Calcul de l'échelle
    If Not EchImpr Then Exit Sub    ' L'échelle n'est pas renseignée
    
    If chkAjuster = 0 Then
      If FichierJournal Then Write #numFichLog, "Impression de l'épure à l'échelle 1/" & txtEchelle
    Else
      If FichierJournal Then Write #numFichLog, "Impression de l'épure ajustée"
    End If
    
    OptGen.Epaisseur = cboEpaisseur
    SaveSetting Appname:=App.Title, SECTION:="Parametres", Key:="Epaisseur ligne", Setting:=OptGen.Epaisseur

     'ECRITURE  de l'en-tête    (unités en cm)
ObjetDessin.CurrentX = 0
ObjetDessin.CurrentY = -1.2
ObjetDessin.FontBold = True
ObjetDessin.FontSize = 12
ObjetDessin.Print Titre1
ObjetDessin.CurrentX = 0
ObjetDessin.CurrentY = -0.7
ObjetDessin.FontSize = 10
ObjetDessin.Print Titre2
      
ObjetDessin.FontBold = False
ObjetDessin.FontSize = 8

frmAnnulImpr.Show

    ' Exclusion de l'entête de la page pour le clipping
    ' facteur fait la conversion cm-->pixel et 2 car l'en-tête fait 2 cm
ExcludeClipRect Printer.hDC, 0, 0, wPixel, facteur * 2
    
' Dessin du cadre
    w = ObjetDessin.ScaleWidth - 0.03       ' 3/10èmm pour pallier un possible effet de bord pour les limites utiles de l'imprimante
    h = ObjetDessin.ScaleHeight - 2.03      ' 2cm pour l'entete + 3/10èmm  """""""""""""""""""""""
    
'    If chkGraph Then
'      ObjetDessin.Line (0, 0)-(w, 0)
'      ObjetDessin.Line (w, 0)-(w, h)
'      ObjetDessin.Line (w, h)-(0, h)
'      ObjetDessin.Line (0, h)-(0, 0)
'    Else
      ObjetDessin.Line (w, h)-(0, 0), , B
'    End If
 ' Dessin de l'épure
    
    On Error GoTo GestErr
    ObjetDessin.DrawWidth = cboEpaisseur
    Rafraichir
    
    If Not frmAnnulImpr.Annul Then Printer.EndDoc
 
End If
  
  Unload Me
  
Exit Sub


GestErr:
  If Err = 102 Then     ' Annulation de l'impression demandée
    If FichierJournal Then Write #numFichLog, "Impression de l'épure abandonnée"
    Printer.KillDoc
    Resume Next
  ElseIf Err = 482 Or Err = 484 Or Err = 486 Then 'Erreur d'imprimante : pas d'imprimante configurée, ou pas allumée, bourrage...482 est un sur-ensemble de 484 et 486
    Unload Me
  Else
    MsgBox IDm_ErrImprim & " (" & Format(Err.Number, "000") & ")" & vbCrLf & Err.Description
    Me.Show vbModal
    Exit Sub
  End If
  Exit Sub
    
End Sub

Private Sub Entete(Titre1 As String, Titre2 As String)
' Détermination des 2 lignes d'en-tête

Titre1 = "GIRATION " & GirationVersion
Titre1 = Titre1 & Space(3) & ID_MaitriseOuvrage
ObjetDessin.FontSize = 10
ObjetDessin.CurrentX = 5
ObjetDessin.CurrentY = -0.7
If chkTitre <> 0 Then
  Titre2 = txtTitre & Space(2)
End If
If chkNomFichiers = 1 Then
  If fCourante.dlgTrajectoire.FileName <> "" Then
    Titre2 = Titre2 & Majus(ID_Trajectoire) & " : " & nomCourt(fCourante.dlgTrajectoire.FileName)
    If fCourante.FDP <> "" Then Titre2 = Titre2 & " - "
  End If
  If fCourante.FDP <> "" Then
    Titre2 = Titre2 & ID_FDP & " : "
    If Extension(fCourante.FDP, vbLowerCase) = "fdp" Then
      Titre2 = Titre2 & nomCourt(fCourante.FDP)
    Else
      Titre2 = Titre2 & fCourante.FDP
    End If
  End If
End If

End Sub

Private Function EchImpr() As Boolean
' Calcul de l'échelle d'impression
  
    EchImpr = True
    
    If chkAjuster = 0 Then
      If txtEchelle = "" Then
        MsgBox ID_Echelle & " " & Idm_Obligatoire, vbExclamation + vbYes
        Me.Show
        txtEchelle.SetFocus
        EchImpr = False
      Else
        Set gtMil = pMilieu(basGauche, hautDroit)    ' le centre de l'écran se positionne au centre de la surface utile de la feuille
        gtFacteurZoom = (1 / txtEchelle) * 100
  '      gtFacteurZoom = (1 / Echelle) * 1000
      
      End If
    End If

End Function

Private Sub cmdOptAff_Click()
        'Options d'affichage

  '        Me.Hide
          frmOptAff.Show vbModal
  '        Me.Show

End Sub

Private Sub chkTitre_Click()

    txtTitre.Enabled = (chkTitre = 1)

End Sub

Public Sub EchelAjustee()

    If flagImpr Then    ' La fonction a déjà été appelée
      With fCourante
        .Milieux.Remove .Milieux.Count                        ' un calcul d'échelle ajustée (fonction calcEchelle) a déjà été fait
        .Echelles.Remove .Echelles.Count
      End With
    End If
    
    With ObjetDessin
      If .ScaleWidth < 0 Or .ScaleHeight < 0 Then ' Bug VB4 16 bits : Scale... sur 2 octets
          MsgBox "Le format de l'imprimante en cours n'est pas géré par GIRATION"
          optNatImpr(1).Enabled = False
          optNatImpr(0) = True
          txtEchelle = ""
          Exit Sub
        Else
          optNatImpr(1).Enabled = True
      End If
      
      .ScaleMode = 3            ' pixel : prépare l'appel de ExcludeClipRect dans cmdOK_Click
      hPixel = .ScaleHeight
      wPixel = .ScaleWidth
      .ScaleMode = 7     ' cm
      facteur = hPixel / .ScaleHeight
'      .ScaleMode = 6     ' mm
          ' paramètres d'origine et d'échelle du dessin
          ' on conserve 2 cm pour les titres
      .ScaleTop = -2
      If .ScaleWidth < 0 Or .ScaleHeight < 0 Then ' Bug VB4 16 bits : Scale... sur 2 octets
        .ScaleWidth = .Width * 0.9 / 567.5        ' 0,9 : marge de sécurité
        .ScaleHeight = .Height * 0.9 / 567.5    ' 567.5 : facteur de conversion Twips --> cm
      End If
      gtOrx = .ScaleWidth / 2
      gtOry = (.ScaleHeight + .ScaleTop) / 2
      
    End With
    If flagImpr Then
      CalcEchelle
    Else  ' 1er appel : recalcul des limites du dessin
      ToutVoir
    End If
    
    memoEchelAjustee = Arrond((1 / gtFacteurZoom) * 100, 0)  ' 100 : conversion m-->cm
'    memoEchelAjustee = arrond((1 / gtFacteurZoom) * 1000, 0)       ' 1000 : conversion m-->mm
'   If chkAjuster.Enabled And chkAjuster = 1 Then Echelle = memoEchelAjustee
   If chkAjuster = 1 Then txtEchelle = memoEchelAjustee

    flagImpr = True
    
End Sub

