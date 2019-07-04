VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options d'importation"
   ClientHeight    =   2910
   ClientLeft      =   1125
   ClientTop       =   1890
   ClientWidth     =   6810
   Icon            =   "IMPORT.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread ListPlans 
      Height          =   2055
      Left            =   240
      OleObjectBlob   =   "IMPORT.frx":000C
      TabIndex        =   4
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   320
      Left            =   5520
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      CausesValidation=   0   'False
      Height          =   320
      Left            =   5520
      TabIndex        =   2
      Top             =   720
      Width           =   1200
   End
   Begin VB.TextBox txtEchelle 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Text            =   "1"
      Top             =   2400
      Width           =   492
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lblUnité 
      Caption         =   "m"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblEchelle 
      Caption         =   "Echelle : 1 unité ="
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmImport   -   IMPORT.FRM
'   Feuille permettant le choix des plans à retenir lors de l'imoort d'un fichier DXF

'**************************************************************************************

Option Explicit

Private sauvEchelle As String

Private Sub cmdCancel_Click()

  Set LecDXF.Graph = Nothing
  Unload Me
  
End Sub

Private Sub txtEchelle_GotFocus()
    sauvEchelle = txtEchelle
End Sub

Private Sub txtEchelle_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True

End Sub

Private Sub txtEchelle_KeyPress(KeyAscii As Integer)
  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False
End Sub

Private Sub txtEchelle_Validate(Cancel As Boolean)
  If controleNumeric(txtEchelle, StrictPositif:=True) Then
    txtEchelle = sauvEchelle
    Cancel = True
  End If
End Sub

Private Sub cmdOK_Click()
Dim Cpt As Integer
Dim nomPlan As String

  If txtEchelle = "" Then
    MsgBox ID_Echelle & " " & Idm_Obligatoire
    txtEchelle.SetFocus
    Exit Sub
  Else
    echDXF = CDbl(txtEchelle)
  End If
  
  With ListPlans
    For Cpt = 1 To .MaxRows
      .Row = Cpt
      If Not .RowHidden Then
        .Col = 2
        nomPlan = .Value
        .Col = 1
        gtPlans(nomPlan).gele = (.Value <> "1")
      End If
    Next
   
  End With
  
  Unload Me

End Sub


Private Sub Form_Load()
    Dim objPlan As PLAN
      
    'Affichage centré de la fenêtre
    'Me.ScaleMode = 1
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
  
  ' Aide contextuelle
  Me.HelpContextID = IMPORTFDP

  With ListPlans
    .MaxRows = 1
    .ColWidth(2) = 30.5
    For Each objPlan In gtPlans
      .Row = .MaxRows
  'Modif AV (19/06/03 :suite à pb sous ACAD2000   (ECRDXF) : Points de définition des cotes de toute façon ignoré de GIRATION
'      If InStr(objPlan.name, "$$$GIR_") = 0 And objPlan.name <> "DEFPOINTS" Then
      If InStr(objPlan.name, "$$$GIR_") = 0 Then
        .Col = 1
        .Value = "1"
        .Col = 2
        .Value = objPlan.name
        .Lock = True
        .Col = 3
        .BackColor = QBColor(tCouleur(Abs(objPlan.couleur)))    ' la couleur peut être négative si plan inactif dans le DXF
        .Col = 4
        .TypeComboBoxCurSel = 0      'type de ligne continu
      Else
        .RowHidden = True
      End If
      .MaxRows = .MaxRows + 1
    Next
  
  .MaxRows = .MaxRows - 1
  
  End With
    
End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

