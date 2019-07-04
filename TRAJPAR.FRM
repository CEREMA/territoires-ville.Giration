VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrajpar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Initialisation de la trajectoire"
   ClientHeight    =   7080
   ClientLeft      =   2865
   ClientTop       =   2985
   ClientWidth     =   5250
   ClipControls    =   0   'False
   Icon            =   "TRAJPAR.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraVehicule 
      Caption         =   " Véhicule "
      Height          =   5415
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3135
      Begin VB.ListBox lstNomVehic 
         Height          =   2790
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtVitesse 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2400
         TabIndex        =   10
         Top             =   4920
         Width           =   375
      End
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Simple"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Articulé"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Bi-articulé"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2880
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblVitesse 
         Caption         =   "Vitesse en Km/h"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   4920
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   320
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Frame fraFDP 
      Caption         =   "Fond de plan"
      Height          =   972
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   4932
      Begin VB.CommandButton cmdOuvreFDP 
         Caption         =   "&Parcourir..."
         Height          =   320
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   1200
      End
      Begin VB.ComboBox cboFdp 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   320
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog dlgFDP 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "fdp"
      DialogTitle     =   "Ouvrir un fond de plan"
   End
End
Attribute VB_Name = "frmTrajpar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmTrajpar   -   TRAJPAR.FRM
'   Saisie des paramètres de la trajectoire : vitesse - Véhicule - Fond de plan éventuel
'   En sortie, ouverture si OK de frmTraject

'**************************************************************************************

'*************              Déclarations            *************
Option Explicit

Private sauvVitesse As Single


Sub cmdOK_Click()
'-----   Bouton OK

Dim FDP As String
    ' remarque: dans le 1er cas, l'appel à Form_Load est déclenché par l'utilisation de frmTraject.Caption
    '                   dans le second, il l'est par l'utilisation de TabForm(NbreFilles).Caption

  
  If lstNomVehic.ListIndex = -1 Then
    MsgBox ID_Véhicule & " " & Idm_Obligatoire
    lstNomVehic.SetFocus
    Exit Sub
  End If
  
  If cboFdp = "<Aucun>" Then cboFdp.ListIndex = -1
  If cboFdp.ListIndex <> -1 Then
    FDP = cboFdp.List(cboFdp.ListIndex)
    If StrComp(nomCourt(FDP) & ".FDP", FDP, vbTextCompare) = 0 Then
      FDP = OptGen.Repert(0) & "\" & FDP      'Fichier récupéré dans la liste des FDP situés dans le répertoire par défaut
      ' sinon   Fichier récupéré par le navigateur
    End If
  End If
  
  creFille nouveau:=True, Nom:=FDP      ' creFille est défini dans GIRATION.BAS
 
End Sub

Private Sub cmdCancel_Click()
' ---- Bouton annuler
    Unload Me

End Sub

Private Sub cmdOuvreFDP_Click()
Dim i%

  With dlgFDP
     .InitDir = gtRepertFDP
     .InitDir = OptGen.Repert(0)
    .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    .Filter = ID_FilterFDPGlobal
    .ShowOpen
    If .FileName <> "" Then
      DoEvents
'      gtRepertFDP = extraiRep(.FileName)
      cboFdp.AddItem .FileName
      With cboFdp
        For i = 0 To .ListCount - 1
          If .List(i) = dlgFDP.FileName Then
            .ListIndex = i
            Exit For
          End If
        Next
      End With
    End If
   End With
   
End Sub

Private Sub Form_Load()
Dim Vehic As VEHICULE, nomfdp As String

'    DégraisserFonte Me

    'Affichage centré de la fenêtre
    Me.ScaleMode = 1
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    txtVitesse = OptGen.VitDef
    
    ' Liste des fonds de plan
    nomfdp = Dir(OptGen.Repert(0) & "\*.FDP")
    If nomfdp <> "" Then cboFdp.AddItem "<Aucun>": cboFdp.ListIndex = 0
    Do While nomfdp <> ""
      cboFdp.AddItem StrConv(nomfdp, vbLowerCase)
      nomfdp = Dir
    Loop
    
  ' Aide contextuelle
  Me.HelpContextID = NOUVTRAJ

End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

Private Sub lstNom_Click()

End Sub

Private Sub optTypVeh_Click(Index As Integer)
Dim i As Integer
      ' Liste des véhicules
  Me.lstNomVehic.Clear
  For i = 1 To ColVéhicules.Count
    If ColVéhicules(i).nbRemorque = Index Then
      lstNomVehic.AddItem Trim(ColVéhicules(i).Nom)
    End If
  Next
End Sub

' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la fonction KeyPress semble en mesure de réafficher la virgule
Private Sub txtVitesse_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True
'    If flagVirgule Then alertVirgule = True
'    If flagVirgule Then KeyCode = 188
'  End If

End Sub

Private Sub txtVitesse_KeyPress(KeyAscii As Integer)
  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False
End Sub

Private Sub txtVitesse_Gotfocus()
  sauvVitesse = txtVitesse
End Sub

Private Sub txtVitesse_Lostfocus()
  
  If controlVitesse Then txtVitesse = sauvVitesse: txtVitesse.SetFocus
  
End Sub

Private Function controlVitesse() As Boolean

  If txtVitesse = "" Then
    MsgBox ID_Vitesse & " " & Idm_Obligatoire
    controlVitesse = True
  ElseIf Not IsNumeric(txtVitesse) Then
      MsgBox IDm_Numerique
      controlVitesse = True
  ElseIf CDbl(txtVitesse) > 50 Or Abs(CDbl(txtVitesse)) < 1 Then
      MsgBox ID_Vitesse & " " & ID_ComprisEntre & " 1 " & ID_Et & " 50 " & ID_UniteVitesse
      controlVitesse = True
  End If
  
End Function

