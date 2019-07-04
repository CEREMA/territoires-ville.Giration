Attribute VB_Name = "strConst"
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97
'     Mise � jour pour la version anglaise : F�vrier 2000
'
'   R�alisation : Andr� VIGNAUD

'   Module standard : strConst   -   STRCONST.BAS

'   Fonctions du module
'     Constantes de chaine � traduire

'**************************************************************************************

Option Explicit

Public Const LGMAXNOMVEHICULE = 16    ' Longueur maximale du nom court du v�hicule (VEHICULE.CLS - LECDXF)-->Export

'**************************************************************************************

' D�claration des constantes de chaine globales
'
'**************************************************************************************

Public Const ID_ExemplDemo = "Exemplaire de d�monstration" ' Vehicule et Apropos
  ' Libell�s des boutons de frmBibV�hicule
Public Const ID_Cr�er = "&Cr�er"                           ' Vehicule et Bibveh
Public Const ID_Visualiser = "&Visualiser"                 ' Vehicule et Bibveh
Public Const ID_Modifier = "&Modifier"                     ' Vehicule et Bibveh
Public Const ID_Fermer = "&Fermer"

Public Const ID_FDP = "Fond de plan"                       ' Imprime (= frmTrajPar.fraFDP)
Public Const ID_EnregFDP = "Enregistrer le " & ID_FDP & " sous"
Public Const ID_ImportFDP = "Importer un " & ID_FDP
Public Const IDm_EnregistrerFDP = "Enregistrer le " & ID_FDP & " associ� �"

Public Const ID_FilterFDP = ID_FDP & " (*.fdp)|*.fdp|"
Public Const ID_AllFiles = "Tout fichier"
Public Const ID_Format = "Format"
Public Const ID_FilterFDPGlobal = ID_FilterFDP & ID_Format & " DXF (*.dxf)|*.dxf|" & ID_AllFiles & " (*.*)|*.*"

Public Const ID_Fichier = "fichier"
Public Const ID_Chemin = "Chemin"
Public Const IDm_Absent = " non trouv�"                   ' blanc initial important
Public Const ID_Imprimante = "imprimante"

Public Const ID_Bibveh = "Biblioth�que de v�hicules"
Public Const ID_V�hicule = "v�hicule"                      ' Vehicule et TrajPar
Public Const Idm_Obligatoire = "obligatoire"               ' Vehicule,Import,TrajPar,Optaff,Imprime
Public Const ID_Vitesse = "Vitesse"                        ' Optgen et TrajPar
Public Const ID_UniteVitesse = "Km/h"                      ' Optgen,TrajPar et Imprime
Public Const ID_seconde = "seconde"                        ' Optgen et TrajPar
Public Const ID_UnitAngle = "Unit� d'angle"                ' Optgen et Traject
Public Const ID_Degres = "Degr�(s)"
Public Const ID_Grades = "Grade(s)"
Public Const ID_Degre = "Degr�"
Public Const ID_Grade = "Grade"
Public Const ID_Echelle = "Echelle"                        ' Import et Imprime
Public Const ID_Trajectoire = "trajectoire"                ' MDIGiration et Imprime

' Fonctions de Zoom, Pan dans frmCadrage et frmTraject
Public Const ID_Premier = "1er"
Public Const ID_Deuxieme = "2�me"
Public Const ID_Dernier = "Dernier"
Public Const ID_Point = "point"
Public Const ID_Coin = "coin"

Public Const IDm_Confirmation = "Confirmation"
Public Const IDm_ConfirmSuppr = "Confirmer la Suppression"

Public Const ID_Et = " et "
Public Const ID_Ou = " ou "

Public Const IDm_Erreur = "Erreur"
Public Const IDm_ErrFatale = "Erreur fatale n� "
Public Const IDm_Anomalie = "Anomalie"

Public Const IDm_ErrImprim = IDm_Erreur & " " & ID_Imprimante           ' MDIGiration et Imprime
Public Const IDm_Numerique = "Num�rique obligatoirement"   ' GirStand,Optgen,TrajPar,Cadrage
Public Const IDm_DifferentDeZero = "Saisir une valeur non nulle"
Public Const IDm_SuperieurAZero = "Saisir une valeur strictement positive"  ' GirStand,Optgen
Public Const IDm_Compris = "Saisir une valeur comprise entre "
Public Const IDm_StrictCompris = "Saisir une valeur strictement comprise entre"

' Libell�s titres tableau de frmTraject : Sert dans la fonction Resol (GirStand.Bas)
Public Const ID_Deplacement = "D�placement"
Public Const ID_DeplacementCourt = "D�pl."
Public Const ID_Rayongir = "Rayon de giration"
Public Const ID_RayongirCourt = "Rayon"
Public Const ID_Anglegir = "Angle de giration"
Public Const ID_AnglegirCourt = "Angle"
Public Const ID_Longueur = "Longueur"
Public Const ID_VitBraq = "Vitesse de braquage"     ' Sert aussi dans Vehicule
Public Const ID_VitBraqCourt = "Vitesse braquage"
Public Const ID_VitBraqTresCourt = "Vit. braq."
Public Const ID_Direction = "Direction"                    '(= lblDirection) - Sert aussi dans frmImprim

Public Const ID_ExportTracteur = "_Tracteur"
Public Const ID_ExportRemorque = "_Remorque1"
Public Const ID_ExportRemorque2 = "_Remorque2"
Public Const ID_ExportRouesAvt = "_Roues_Avt"
Public Const ID_ExportRouesArr = "_Roues_Arr"
Public Const ID_ExportRouesRmq1 = "_Roues_Rmq1"
Public Const ID_ExportRouesRmq2 = "_Roues_Rmq2"

'v3.3 : Suppression de la protection suite au remplacement de CopyControl par CopyMinder
' Protection : CopyControl et lireProtect
'Public Const ID_GestionProtection = "Gestion de la Protection"

'**************************************************************************************
' Module GirationMain
'**************************************************************************************

Public Const IDm_IncompatiblBibvehVersiondemo = "Biblioth�que de v�hicules non utilisable par la version de d�monstration"
Public Const IDm_InitInterrupt = "Initialisation de GIRATION interrompue"
Public Const IDm_DroitsBibVeh = "Droits d'acc�s � la biblioth�que de v�hicules insuffisants"

'**************************************************************************************
' Module GraphStand
'**************************************************************************************

Public Const IDm_AnomalieDessin = "Anomalie dans le dessin - " & IDm_Erreur & " "

'**************************************************************************************
' Module LecDXF
'**************************************************************************************

Public Const ID_LIGNE = "ligne"
Public Const ID_Code = "Code"
Public Const ID_Attendu = "attendu"
Public Const ID_NombreEntier = "nombre entier"

Public Const IDm_Incorrect = "incorrect"
Public Const IDm_DXFIncorrect = ID_Code & " DXF " & IDm_Incorrect
Public Const ID_DXFVersion = "Version DXF"
Public Const ID_NonGeree = " non g�r�e par " ' blancs encadrant essentiels
Public Const ID_Lecture = "Lecture"
Public Const ID_LectureFichier = ID_Lecture & " - " & ID_Fichier & " "
Public Const ID_Plan = "Plan"
Public Const IDm_AbsentTablePlan = " absent de la table des plans"
Public Const IDm_EntiteSansPlan = "Pas de plan d�fini pour cette entit�"
Public Const IDm_UnSeulPointPline = "Une polyligne doit comporter au moins 2 points"

Public Const ID_RechercheLimites = "Recherche des limites..."
Public Const IDm_Err101 = "Aucun �l�ment interpr�table par GIRATION n'a �t� trouv� dans "
Public Const IDm_Err103 = ID_Fichier & IDm_Absent
Public Const IDm_FDPRefus� = ID_FDP & " non charg�"
Public Const IDm_FinPrematuree = "Fin pr�matur�e atteinte"

'**************************************************************************************
' Module CopyControl
'**************************************************************************************
Public Const IDm_ProduitAbsent = "Produit non install�"
Public Const IDm_NumLicence = "Le num�ro de licence ne correspond pas"
Public Const IDm_Jeton = "Jeton introuvable"
Public Const IDm_DisqueProt�g� = "V�rification impossible : le disque est prot�g� en �criture"
Public Const IDm_DisquetteProt�g�e = "V�rification impossible : la disquette est prot�g�e en �criture"
Public Const IDm_TropDUtilisateurs = "Veuillez recommencer plus tard" & vbCrLf & "Trop d'utilisateurs sont pr�sents"
Public Const IDm_GestionLicence = "La licence doit �tre activ�e" & vbCrLf & "Lancer le programme Licence.exe gr�ce au bouton DEMARRER, puis menu PROGRAMME / " & "GIRATION" & " / GESTION LICENCE"
Public Const IDm_ProtectionAbsente = "GIRATION n'a pas trouv� la protection"

'**************************************************************************************
' Module frmAPropos
'**************************************************************************************
'Public Const ID_Licence = "Licence n�"
Public Const ID_Licence = "Release n�"

'**************************************************************************************
' Module frmBibV�hicule
'**************************************************************************************
Public Const ID_Supprimer = "Supprimer"

'**************************************************************************************
' Module frmCadrage
'**************************************************************************************
Public Const IDm_SeuilZoom = "Seuil de Zoom atteint"
Public Const IDm_Invalid = "Invalide"
Public Const ID_ClicContextuel = "Clic droit pour menu contextuel - Echap pour sortir"

'**************************************************************************************
' Module frmImporBib
'**************************************************************************************
Public Const IDm_NomVehicUtilise = "Nom de v�hicule d�j� utilis�"

'**************************************************************************************
' Module frmImprim
'**************************************************************************************
Public Const ID_ImprimanteEnCours = "Imprimante en cours"
Public Const ID_TitreImpressionVersionDemo = ID_ExemplDemo & " - non utilisable pour un projet"
Public Const ID_MaitriseOuvrage = "CERTU - CETE de l'Ouest"

'**************************************************************************************
' Module frmLargeur
'**************************************************************************************
Public Const IDm_DebordementMini = "Saisir un d�bordement strictement sup�rieur �"
Public Const IDm_SurlargeurMini = "Saisir une surlargeur strictement sup�rieure �"

'**************************************************************************************
' Module frmOptAff
'**************************************************************************************
Public Const ID_Saisie = "Saisie"

'**************************************************************************************
' Module frmOptGen
'**************************************************************************************
Public Const IDm_CheminAbsent = ID_Chemin & IDm_Absent
Public Const IDm_MaxVitesse = "La vitesse ne doit pas d�passer"
Public Const IDm_NomBibvehInvalide = "Nom de biblioth�que non valide"
Public Const IDm_Absente = " absente" ' s'accorde avec biblioth�que de v�hicules blanc initial
Public Const IDm_CreationDossierRefusee = "Cr�ation de dossier refus�e"
Public Const IDm_CreerDossier = "Cr�er le dossier " '(conserver le blanc en fin de chaine)

'**************************************************************************************
' Module frmTrajpar
'**************************************************************************************
Public Const ID_ComprisEntre = "comprise entre"

'**************************************************************************************
' Module frmV�hicule
'**************************************************************************************
Public Const ID_Simple = "Simple"
Public Const ID_Articul� = "Articul�"
Public Const ID_BiArticul� = "Bi-articul�"

'**************************************************************************************
' Module MDIGiration
'**************************************************************************************
Public Const IDm_BibVehVide = "La biblioth�que de v�hicules est vide"
Public Const IDm_RemplacerFDP = "La trajectoire comporte d�j� un " & ID_FDP & vbCrLf & "Remplacer"
Public Const IDm_MRUFichierDisparu = "Fichier introuvable" & vbCrLf & vbCrLf & "Il doit avoir �t� effac� ou chang� de dossier"

