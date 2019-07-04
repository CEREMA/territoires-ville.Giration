VERSION 5.00
Begin VB.Form frmBibVéhicule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bibliothèque de véhicules"
   ClientHeight    =   4860
   ClientLeft      =   2310
   ClientTop       =   4200
   ClientWidth     =   5460
   ClipControls    =   0   'False
   Icon            =   "BIBVEH.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstNom 
      Height          =   2790
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Frame fraVehicule 
      Caption         =   " Véhicule "
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3132
      Begin VB.OptionButton optTypVeh 
         Caption         =   "Bi-articulé"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1335
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
         Caption         =   "Simple"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox picBoutons 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4860
      Left            =   3885
      ScaleHeight     =   4860
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Aide"
         Height          =   320
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Créer..."
         Height          =   320
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTerminer 
         Cancel          =   -1  'True
         Caption         =   "&Terminer"
         Height          =   320
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Supprimer"
         Enabled         =   0   'False
         Height          =   320
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Modifier..."
         Enabled         =   0   'False
         Height          =   320
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmBibVéhicule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmBibVéhicule   -   BIBVEH.FRM
'   Feuille d'accueil de la Bibliothèque de Véhicules (frmVéhicul)

'**************************************************************************************
Option Explicit
Private numFichVeh As Integer
Public Modif As Boolean
Private DrapeauBibliVide As Boolean
Private mVéhicule As VEHICULE
Private mTypeVeh As String ' "S", "A", ou "V" pour Simple, Articulé, Bi-articulé
Public ChargementEnCours As Boolean

Private Sub ActiverBoutons()
  
  If mVéhicule Is Nothing Then
    Me.cmdAction(1).Enabled = False
    Me.cmdAction(2).Enabled = False
  Else
    Me.cmdAction(1).Enabled = True
    Me.cmdAction(2).Enabled = Not mVéhicule.Protégé
    If EnVisu Then
      Me.cmdAction(1).Caption = ID_Visualiser
      Me.cmdAction(2).Enabled = False
    Else
      Me.cmdAction(1).Caption = ID_Modifier
      Me.cmdAction(2).Enabled = True
    End If
  End If
End Sub

Public Function EnVisu() As Boolean

  If gtBibliVerrouillée Then
    EnVisu = True
  Else
    If mVéhicule Is Nothing Then
      EnVisu = True
    Else
      EnVisu = mVéhicule.Protégé
    End If
  End If
  
End Function

'********************************************************************************
'Choix d'une action :  Créer, Modifier, Supprimer un véhicule
'********************************************************************************
Private Sub cmdAction_Click(Index As Integer)
Dim LectureSeule As Boolean
Dim OuvrirVeh As Boolean
Dim unVéhicule As VEHICULE

  Select Case Index
  Case 0
    Set unVéhicule = DemandeNomVéhic
    If Not unVéhicule Is Nothing Then
      Set mVéhicule = unVéhicule
      OuvrirVeh = True
      Me.optTypVeh(unVéhicule.nbRemorque) = True
    End If
    
  Case 1

    OuvrirVeh = True
  
  Case 2
    'Supprimer un véhicule
    If MsgBox(ID_Supprimer & " " & mVéhicule.Nom & " ?", vbYesNo + vbExclamation + vbDefaultButton2, IDm_ConfirmSuppr) = vbYes Then
      SupprimerVéhicule
      Set mVéhicule = Nothing
      mTypeVeh = ""
      ActiverBoutons
    End If

  End Select
  
  If OuvrirVeh Then
  'Créer,Modifier ou Visualiser un véhicule
    If Not EnVisu Then
      'Vérifier l'autorisation de mise à jour
      Close numFichVeh
      If BibliEnMiseAJour Then
        LectureSeule = True
      Else
      ' Verrouilage de la bibliothèque en écriture
        numFichVeh = FreeFile
        Open GirationVeh For Random Lock Write As numFichVeh
      End If
    End If
    
    If Not LectureSeule Then
      With frmVéhicule
        If .contextBibliVeh Then  ' Bibliothèque en cours d'utilisation : la feuille est simplement cachée
          Unload frmVéhicule
        Else
          .contextBibliVeh = True
        End If
        .Show vbModal
        If Not ExistVeh(mVéhicule.Nom, ColVéhicules) Then
          'abandon de la création
          Set mVéhicule = Nothing
          mTypeVeh = ""
        End If
      End With
    End If
  End If
  
End Sub

Public Sub AjouterVéhicule(ByVal unVéhicule As VEHICULE)

On Error GoTo GestErr

  If Not unVéhicule Is Nothing Then
    Set mVéhicule = unVéhicule
    optTypVeh_Click Numopt(optTypVeh)
  End If
Exit Sub
GestErr:
  If Err <> ErrVéhicule Then
    ErreurNonFatale "BibVéhicule.Ajouter"
  End If
  Err.Raise ErrVéhicule
End Sub

Public Sub RenommerVéhicule(ByVal unVéhicule As VEHICULE)
  
On Error GoTo GestErr
  
  If Not unVéhicule Is Nothing Then
    SupprimerVéhicule
    AjouterVéhicule unVéhicule
  End If

Exit Sub
GestErr:
  If Err <> ErrVéhicule Then
    ErreurNonFatale "BibVéhicule.Renommer"
  End If
  Err.Raise ErrVéhicule
End Sub

Private Sub SupprimerVéhicule()
  ColVéhicules.Remove UCase(mVéhicule.Nom)
  lstNom.RemoveItem Me.lstNom.ListIndex
  If Not VersionDemo Then Modif = True
End Sub

Private Function DemandeNomVéhic() As VEHICULE
Dim NumOptTyp As Integer

  With frmNewVehicule
    NumOptTyp = Numopt(optTypVeh)
    If NumOptTyp <> -1 Then
      .optTypVeh(NumOptTyp).Value = True
    End If
    .Show vbModal, Me
    If .ValidOK Then
      .ValidOK = False
      
      Set DemandeNomVéhic = New VEHICULE
      DemandeNomVéhic.Nom = Trim(.txtNom.Text)
      Select Case Numopt(.optTypVeh)
      Case 0
        mTypeVeh = Left(ID_Simple, 1)
        DemandeNomVéhic.nbRemorque = 0
      Case 1
        mTypeVeh = Left(ID_Articulé, 1)
        DemandeNomVéhic.nbRemorque = 1
      Case 2
        mTypeVeh = Left(ID_BiArticulé, 1)
        DemandeNomVéhic.nbRemorque = 2
      End Select
    End If
  End With
  
  Unload frmNewVehicule
  
End Function

Private Sub cmdTerminer_Click()

Dim VehicTab As StructVéhicule
Dim Vehic As VEHICULE
Dim i As Integer

  Unload Me   ' Décharge également frmVéhicule éventuellement cachée

  If Modif Then  ' Enregistrement de la bibliothèque
    Modif = False
    
    If Not BibliEnMiseAJour Then
      If ExistFich(GirationVeh) Then Kill GirationVeh
      numFichVeh = FreeFile
      Open GirationVeh For Random Lock Read Write As numFichVeh Len = Len(VehicTab)
      Put #numFichVeh, 1, vehicVersion
        For Each Vehic In ColVéhicules
          'Mise au format StructVéhicule l'objet Véhicule
          ConvObjetStruct Vehic, VehicTab
          'Ecriture du véhicule dans le fichier
          i = i + 1
          Put #numFichVeh, i + 1, VehicTab
        Next
      Close numFichVeh
    End If

  End If

End Sub

Private Sub Form_Activate()
  If ChargementEnCours Then
    ChargementEnCours = False
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Modif And UnloadMode = vbFormControlMenu Then
    Cancel = (MsgBox("Abandonner les modifications de la bibliothèque" & vbCrLf & vbCrLf, vbYesNo + vbQuestion + vbDefaultButton2, IDm_Confirmation)) = vbNo
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

      If frmVéhicule.contextBibliVeh Then ' La feuille frmVéhicule est encore chargée
        Unload frmVéhicule
        frmVéhicule.contextBibliVeh = False
      End If
      Close numFichVeh
      gtBibliVerrouillée = False
      If DrapeauBibliVide And ColVéhicules.Count = 0 Then Kill GirationVeh

End Sub

Public Sub VehDelete()
  cmdAction(2) = True
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Vehic As VEHICULE

On Error GoTo GestErr
     
  ChargementEnCours = True
  
  If BibliEnMiseAJour(premier:=True) Then
    Me.cmdAction(0).Enabled = False
    
  Else
    ' Verrouilage de la bibliothèque en écriture
    DrapeauBibliVide = Not ExistFich(GirationVeh)
    numFichVeh = FreeFile
    Open GirationVeh For Random Lock Write As numFichVeh
  End If

  'Affichage centré de la fenêtre
  Me.ScaleMode = vbTwips
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2

'  optTypVeh(0) = True

   ' Aide en ligne contexte
   Me.HelpContextID = BIBVEHICUL

Exit Sub

GestErr:
    ErrGeneral

End Sub

Private Sub cmdHelp_Click()
  
  SendKeys "{F1}", True
  
End Sub

Public Property Get Véhicule() As VEHICULE
  Set Véhicule = mVéhicule
End Property

Public Property Let Véhicule(ByVal vNewValue As VEHICULE)
  Set mVéhicule = vNewValue
End Property


Private Sub lstNom_Click()
Dim i As Integer

  Set mVéhicule = ColVéhicules(UCase(lstNom.Text))

  cmdAction(1).Enabled = True
  With mVéhicule
    If .Protégé Or gtBibliVerrouillée Then 'Or .nbUtilisation > 0 Then
      cmdAction(1).Caption = ID_Visualiser & "..."
      cmdAction(2).Enabled = False
    Else
      cmdAction(1).Caption = ID_Modifier & "..."
      cmdAction(2).Enabled = True
    End If
 
    optTypVeh(.nbRemorque) = True
  End With
  
  ' Gérer l'activité des boutons en fonction du choix
  ActiverBoutons

End Sub

Private Sub lstNom_DblClick()
  If cmdAction(1).Enabled Then cmdAction(1).Value = True
End Sub

Private Sub optTypVeh_Click(Index As Integer)
Dim i As Integer

      ' Liste des véhicules
    With lstNom
      Me.lstNom.Clear
      For i = 1 To ColVéhicules.Count
        If ColVéhicules(i).nbRemorque = Index Then
          .AddItem Trim(ColVéhicules(i).Nom)
          If Not mVéhicule Is Nothing Then
            If mVéhicule.Nom = Trim(ColVéhicules(i).Nom) Then
              .ListIndex = .ListCount - 1
            End If
          End If
        End If
      Next
      
    End With

End Sub
