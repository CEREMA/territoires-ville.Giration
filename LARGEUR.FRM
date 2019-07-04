VERSION 5.00
Begin VB.Form frmLargeur 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surlargeur"
   ClientHeight    =   4395
   ClientLeft      =   3330
   ClientTop       =   2460
   ClientWidth     =   4290
   Icon            =   "LARGEUR.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   320
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Frame fraDebordement 
      Caption         =   "Débordements"
      Height          =   1572
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   2652
      Begin VB.TextBox txtDebord 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtDebord 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblDebArr 
         Caption         =   "Arrière"
         Height          =   288
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   612
      End
      Begin VB.Label lblDebAvant 
         Caption         =   "Avant"
         Height          =   288
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.Frame fraSurLargeur 
      Caption         =   " Surlargeur(s)"
      Height          =   2052
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2652
      Begin VB.TextBox txtSurLarg 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtSurLarg 
         Enabled         =   0   'False
         Height          =   288
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   624
      End
      Begin VB.TextBox txtSurLarg 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblSurLargeur 
         Caption         =   "Tracteur"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblSurLargeur 
         Caption         =   "Remorque 1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSurLargeur 
         Caption         =   "Remorque 2"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      CausesValidation=   0   'False
      Height          =   320
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   1200
   End
End
Attribute VB_Name = "frmLargeur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmLargeur   -   LARGEUR.FRM
'   Feuille permettant de mettre des surlargeurs et des débordements
'   à un véhicule de la bibliothèque, sans modifier celle-ci, lors d'une étude de trajectoire particulière

'**************************************************************************************
Option Explicit

Const AVANT = 0, ARRIERE = 1

Private sauvDebord(1), sauvSurLarg(2) As String, MinSurLarg(2) As Double, MinDebord(2) As Double


Private Sub cmdCancel_Click()

      If FichierJournal Then Write #numFichLog, "Surlageur Annuler"
  
  Unload Me
  
End Sub

Private Sub cmdOK_Click()
Dim numRemorque As Integer

      If FichierJournal Then Write #numFichLog, "Surlageur OK"

' L'utilisateur a fait RETOUR-CHARIOT
If ActiveControl.name <> "cmdOK" Then SendKeys "{TAB}": Exit Sub

  With fCourante
    .Debord(AVANT) = recup(txtDebord(AVANT))
    .Debord(ARRIERE) = recup(txtDebord(ARRIERE))
    For numRemorque = 0 To .nbRemorque
      .SurLarg(numRemorque) = recup(txtSurLarg(numRemorque))
    Next
  End With
  
  Me.Hide
  ' Recrée les éléments de carrosserie avec les nouvelles dimensions
  fCourante.CreCarross surLargeur:=True
        
    If FichierJournal Then Write #numFichLog, "Surlageur terminée"

  Rafraichir

Unload Me

End Sub

' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la focntion KeyPress semble en mesure de réafficher la virgule
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyDecimal And Shift = 0 Then alertVirgule = True
'    If flagVirgule Then alertVirgule = True
'    If flagVirgule Then KeyCode = 188
'  End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If alertVirgule Then KeyAscii = gtPtDecimal: alertVirgule = False
End Sub

Private Sub Form_Load()
Dim numRemorque As Integer
Dim nbRemorque As Integer

      If FichierJournal Then Write #numFichLog, "Ouverture de surlageur"
    
'      DégraisserFonte Me
      
    'Affichage centré de la fenêtre
      Me.ScaleMode = 1
      Me.Left = (Screen.Width - Me.Width) / 2
      Me.Top = (Screen.Height - Me.Height) / 2
      
    ' Aide contextuelle
    Me.HelpContextID = LARGEUR
    
  With fCourante
    nbRemorque = .nbRemorque
    txtDebord(AVANT) = affiche(.Debord(AVANT))
    txtDebord(ARRIERE) = affiche(.Debord(ARRIERE))
    With .Vehic
      MinDebord(AVANT) = (.E(0) + .PAFav(0) + .PAFar(0)) / 2
      MinDebord(ARRIERE) = (.E(nbRemorque) + .PAFav(nbRemorque) + .PAFar(nbRemorque)) / 2
    End With
    For numRemorque = 0 To .nbRemorque
      lblSurLargeur(numRemorque).Enabled = True
      txtSurLarg(numRemorque).Enabled = True
      txtSurLarg(numRemorque) = affiche(.SurLarg(numRemorque))
      MinSurLarg(numRemorque) = .Vehic.l(numRemorque) / 2
    Next
  End With  ' fCourante
  
  
End Sub

Private Sub cmdHelp_Click()
  SendKeys "{F1}", True
End Sub

Private Function recup(v As String) As Double
  If v = "" Then recup = 0# Else recup = CDbl(v)
End Function

Private Function affiche(ByVal v As Variant, Optional flag As Variant) As Variant
  If v = "" Then affiche = "": Exit Function
  If v = 0# And IsMissing(flag) Then affiche = "" Else affiche = Format(v, "####0.000")
End Function


Private Sub txtDebord_GotFocus(Index As Integer)
  sauvDebord(Index) = txtDebord(Index)
End Sub

Private Sub txtDebord_Validate(Index As Integer, Cancel As Boolean)

  If txtDebord(Index) = "" Then
  ElseIf controleNumeric(txtDebord(Index), StrictPositif:=False) Then
    Cancel = True
  ElseIf txtDebord(Index) <= -MinDebord(Index) Then
    MsgBox IDm_DebordementMini & " " & Format(-MinDebord(Index), "0.0##") & " m"
    Cancel = True
  End If
  
  If Cancel Then
    txtDebord(Index) = sauvDebord(Index)
  Else
    txtDebord(Index) = affiche(txtDebord(Index))
  End If

End Sub

Private Sub txtSurLarg_GotFocus(Index As Integer)
  sauvSurLarg(Index) = txtSurLarg(Index)
End Sub

Private Sub txtSurLarg_Validate(Index As Integer, Cancel As Boolean)
  If txtSurLarg(Index) = "" Then
  ElseIf controleNumeric(txtSurLarg(Index), StrictPositif:=False) Then
    Cancel = True
  ElseIf txtSurLarg(Index) <= -MinSurLarg(Index) Then
    MsgBox IDm_SurlargeurMini & " " & Format(-MinSurLarg(Index), "0.0##") & " m"
    Cancel = True
  End If
  
  If Cancel Then
    txtSurLarg(Index) = sauvSurLarg(Index)
  Else
    txtSurLarg(Index) = affiche(txtSurLarg(Index))
  End If

End Sub
