#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <Json.au3>
#include <FileConstants.au3> 

;==========================================================================================================================
;
; Nom du fichier: ImpressionAutoSage100.au3
; Auteur: Baillif Etienne
; Date de création: 07/04/2023
; Modificaiton: 05/06/2023
; Version: v1.03
; Description: Script pour automatiser l'impression de documents de bon de livraison dans Sage100
;
;==========================================================================================================================


; Récupère le premier argument de ligne de commande comme chemin de fichier
$filePath = $CmdLine[1]		;Chemin de l'éxécutable de sage100 Gescom
$fileDbPath = $CmdLine[2]   ;Chemin du fichier pointeur .gcm
$jsonList = $CmdLine[3] 	;List json qui contient les numéro de BL a imprimer
$dbName = $CmdLine[4]   	;Nom de la bdd
$docArray = Json_Decode($jsonList)

Global $LockFile = @ScriptDir & "\AutoItLockFile.txt"  ;répertoire actuelle du script
Global $LogFile = @ScriptDir &  "\autoItLog.txt"       ;Log
$actualWindowName = "Sage 100cloud Gestion commerciale Premium" 

;créer le fichier log si il n'existe pas
WriteToLog($LogFile, "")

;**************************************************Vérification de l'état du script autoit*****************************************************
; Vérifie si le fichier de verrouillage existe
If FileExists($LockFile) Then
	; Si le fichier existe, cela signifie qu'un autre script est en cours d'exécution, donc on arrête le script actuel
	WriteToLog($LogFile,"un script est déja en cours d'éxécution, fermeture...")
	Exit(0)
EndIf

WriteToLog($LogFile,"le fichier lock existe pas création du fichier" & $LockFile)
 ; Crée le fichier de verrouillage pour indiquer que le script est en cours d'exécution
Local $hFile = FileOpen($LockFile, 2)
WriteToLog($LogFile,"le script est en cours d'exécution")
FileClose($hFile)

;***********************************************************************************************************************************************************

;**************************************************Vérification de l'état de sage 100***********************************************************************

;récupérer la poignée de la fenetre
$hwndAccueil = GetHandle("GeCoMaes") 
WinActivate ($hwndAccueil)
If @error Then
	WriteToLog($logFile,"La fenetre d'accueil n'as pas été trouvée")
EndIf

; Vérifier si la fenêtre Sage vierge existe
If WinExists($hwndAccueil) Then
	
    WriteToLog($LogFile, "une fenetre sage 100 est déja ouverte avec la poignée : " & $hwndAccueil)  
Else
	WriteToLog($LogFile, "Sage 100 n'est pas lancé")
    ;Lance Sage100
    Run($filePath)
    WriteToLog($LogFile, "Lancement de l'exe " & $filePath)
	Sleep(3000)
	$hwndAccueil = GetHandle("GeCoMaes") 
	WinActivate ($hwndAccueil)
	If @error Then
		WriteToLog($logFile,"Arret du script: la fenetre d'accueil n'as pas été trouvé")
		EndScript(0)
	EndIf
EndIf

; Connexion à la base de données
OuvrirFichierGcm()

;***********************************************************************************************************************************************************

;**************************************************Gestion du message bloquant sur la fermeture fiscal******************************************************
WriteToLog($LogFile,"gestion du message de fermeture fiscal")
Sleep(2000)
$hWndFermetureFiscal = GetHandle("#32770")
If $hWndFermetureFiscal = False Then  
	; soit le msg fiscal n'est pas apparu, soit le handle n'a pas pu etre récupérer
	WriteToLog($logFile,"La fenetre de fermeture fiscal n'as pas été trouvée ou n'as pas été détéctée")
	; dans le cas la fenetre existe mais que le programme n'as pas détécté on utilise le handleaccueil pour faire la cmd echap 
	WriteToLog($LogFile,"Tentative de fermeture du msg fiscal avec échap avec focus sur" & $hwndAccueil)
	WinActivate ($hwndAccueil)              
		If WinWaitActive($hwndAccueil,"",10) Then  	
			Sleep(4000)
			WinActivate ($hwndAccueil)          
			WriteToLog($LogFile,"Appuie sur ESC pour fermer fenetre bloquante")
			Send("{ESC}")
		Else 
			WriteToLog($LogFile,"La fenêtre" & $hwndAccueil & "n'a pas été détectée après 10 secondes, fermeture...")
			WinClose($hwndAccueil)
			EndScript(0)
		EndIf
Else	
	; cas ou la poigné a été trouvé donc le message est bien actif
	WriteToLog($LogFile,"fenetre de fermeture fiscal actif")
	If WinWaitActive($hWndFermetureFiscal,"",10) Then  
		WinActivate ($hWndFermetureFiscal)  
		If ControlClick($hWndFermetureFiscal, "OK", "Button1") = 1 Then
			WriteToLog($logFile, "succes de la fermeture du msg en cliquant sur OK")
		Else  ;OU ALORS FERMER CARREMENT LE PROGRAMME
			WriteToLog($logFile, "echec de la fermeture du msg en essayant de  cliquer sur OK")
			WinActivate ($hwndAccueil)          
			WriteToLog($LogFile,"tentative de fermeture du msg fiscal avec échap avec focus sur" & $hwndAccueil)
			Send("{ESC}")
			Sleep(2000)
			Send("{ESC}")
		EndIf
	EndIf
EndIf
;***********************************************************************************************************************************************************


;**************************************************PROCESS IMPRESSION***************************************************************************************

WinActivate($hwndAccueil)
If WinWaitActive($hwndAccueil,"",15) Then
		
; Attendre que la liste déroulante soit affichée
Sleep(1000)
WriteToLog($LogFile,"Processus d'impression...") 

;Boucle sur tous les bon de livraison (récupéré depuis le programme C#)
For $i = 0 To UBound($docArray) - 1
	
	;**************************************************Ouverture de la fenetre "Liste des documents de vente"******************************************************
	WriteToLog($LogFile,"Tentative d'ouverture de la fenetre liste des documents de vente")
	WinActivate($hwndAccueil)
	WriteToLog($LogFile,"En attente de la fenetre "& $hwndAccueil& "...")
	If WinWaitActive($hwndAccueil, "", 10) Then
		WinActivate($hwndAccueil) 
		Send("!t")
		WriteToLog($LogFile,"cmd alt+t")
		; Sélectionner l'item Documents de ventes
		Send("v")
		WriteToLog($LogFile,"cmd alt+v")
		Send("{ENTER}")	
	Else 
		WriteToLog($LogFile,"La fenêtre" & $hwndAccueil & "n'a pas été détectée après 10 secondes, fermeture...")
		WinClose($hwndAccueil)
		EndScript(0)
	EndIf

	;***********************************************************************************************************************************************************
	WinActivate($hwndAccueil)
	WriteToLog($LogFile,"***************Impression du document n°"& 0 & ": " & $docArray[$i] &"***************")
	; Cliquer sur le bouton "Actions"
	;récupérer le handle depuis le controle de la fenetre liste des documents
	;$hWndDocVente = ControlGetHandle($hwndAccueil,"", "[CLASS:view.menubutton]")
		;If @error Then
		;WriteToLog($logFile, "Une erreur est survenue lors de la récupération du handle de la fenêtre documentVente")
	;EndIf
	ControlClick($hwndAccueil, "Actions", "view.menubutton5")    ;5 sur l'environnement test
	WriteToLog($LogFile,"cmd btn Action")
	Sleep(2000)
	ControlSend($hwndAccueil,"Actions","[CLASS:view.menubutton]","{DOWN 5}")
	WriteToLog($LogFile,"cmd down x5")
	ControlSend($hwndAccueil,"Actions","[CLASS:view.menubutton]","{ENTER}")


	;***************Fenetre pour saisir les numéros des bons de livraisons;***************
	WriteToLog($LogFile,"wait fenetre Impression Liste de documents... ")
	Sleep(2000) ;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	;récupérér la poignet du la fenetre Impression Liste de documents
	$hwndImpression = GetHandle("view.dialogwnd")
	if $hwndImpression = False Then
		Send("^w")
		WinClose($hwndAccueil)
		EndScript(0)
	EndIf


	;EtapeDeVerification($hwndImpression)  ;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	if WinWait($hwndImpression,"",10) = 0 Then
		WriteToLog($LogFile,"La feneêtre d'impression ne s'est pas ouverte après 10s")
		Send("^w")
		WinClose($hwndAccueil)
		EndScript(0)
	EndIf
	
	
	WriteToLog($LogFile,"la fenetre Impression Liste de documents existe! ")
	
	Local $hWnd = WinWait("[CLASS:view.dialogwnd]", "", 10)
	
	;selectionner "document" dans la fenetre d'impression dans la combobox "type état"
	ControlClick($hWnd, "Liste", "ComboBox2") ;2 sur l'envr prod
	Sleep(500)
	ControlSend($hWnd,"Liste", "ComboBox2","{DOWN 2}") ;2 sur l'envr prod
	;ControlSend($hWnd,"Liste","ComboBox2","{ENTER}") ;2 sur l'envr prod
	Send("{ENTER}")
	
	; selection de l'option Bon de livraison dans le champs "Document"
    ControlClick($hWnd, "Tous", "ComboBox3")	
	Sleep(500)
	ControlSend($hWnd,"Tous", "ComboBox3","{DOWN 4}")
	;ControlSend("Impression Liste de documents","Tous", "ComboBox3","{ENTER}")
	Send("{ENTER}")
	Sleep(500)
	
		;***************saisie des numéros des bons de livraisons;***************
	ControlClick($hWnd, "", "Edit3")
	ControlSend($hWnd, "", "Edit3", $docArray[$i])
	WriteToLog($LogFile,"saisie du numéro de document num:" & $docArray[$i])
	Sleep(500)
	ControlClick($hWnd, "", "Edit4")	
	ControlSend($hWnd, "", "Edit4", $docArray[$i])
	Sleep(500)

	EtapeDeVerification($hWnd)

	ControlClick($hWnd, "OK", "Button1")
	Sleep(4000)

	;Dans la cas ou la fenetre de selection de modele s'ouvre
	If WinExists("Sélectionner le modèle", "") Then 
		WriteToLog($LogFile,"La fenetre de selection de modele existe")
		WriteToLog($LogFile,"Le client lié au BL n'as pas de modele d'impression, fermeture...")
		ControlSend("Sélectionner le modèle", "", "[CLASS:#32770]", "{ESC}")
		Send("{ESC}")
		WriteToLog($LogFile,"cmd Annuler")
	Else
		;Boutton d'impression
			$hwndPrinter =  GetHandle("#32770")  
			if($hwndPrinter = False) Then
				Send("{ESC}")
				Send("{ESC}")
				Send("{ESC}")
				Send("{ESC}")
				WinClose($hwndAccueil)
				EndScript(0)
			EndIf
		if WinWaitActive($hwndPrinter, "", 10) then
		ControlClick($hwndPrinter, "OK", "Button19")
		;WinClose($hwndPrinter) ;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		;Sleep(5000)
		WriteToLog($LogFile,"impression de :" & $docArray[$i])
		EndIf
		
		if WinWaitActive($hwndPrinter, "", 5) then 
			Send("{ESC}") 
		EndIf
	EndIf
	Sleep(3000)
	Send("{ESC}")
	Send("{ESC}")
	Send("{ESC}")
	Send("{ESC}")
Next

;******************************************************************Fin impression*********************************************************************

WriteToLog($LogFile,"fin d'impression")

;ferme la base sage 100
Send("{ESC}")
Send("{ESC}")
Send("{ESC}")
Send("{ESC}")
Send("^w")
EndScript(1) ; ferme le programme et retourne 1 au code c# pour confirmé la fin d"exécution du script

Else
	WinClose($hwndAccueil)
	WriteToLog($LogFile,"La fenetre "&$dbName&" - Sage 100 Gestion commerciale Premium n'a pas été active apres 20s")
    EndScript(0)
EndIf

;******************************************************************FIN SCRIPT*********************************************************************

;FONCTIONS...

Func OuvrirFichierGcm()
	;**************************************************Ouverture du fichier commercial .gcm dans sage100******************************************************
	
	WriteToLog($LogFile," tentative d'ouverture du fichier commercial .gcm")
	Sleep(3000)
	WriteToLog($LogFile,"Poignée avant : " & $hwndAccueil)
	;$hwndAccueil = GetHandle("GeCoMaes") 
	;WriteToLog($LogFile,"Poignée après : " & $hwndAccueil)
	;WinSetOnTop($hwndAccueil)
	WinActivate ($hwndAccueil)
	ControlClick($hwndAccueil,"","[CLASS:view.apptitle]")
	If WinActive($hwndAccueil) Then
		WriteToLog($LogFile,"La fenêtre " & $hwndAccueil & " est active.")
	Else
		WriteToLog($LogFile,"La fenêtre " & $hwndAccueil & " n'est pas active.")
		ControlClick("[CLASS:GeCoMaes]","","[CLASS:view.apptitle]")
		WinActivate ("[CLASS:GeCoMaes]")
			If WinActive("[CLASS:GeCoMaes]") Then
				WriteToLog($LogFile,"La fenêtre " & $hwndAccueil & " est active.")
			Else
				WriteToLog($LogFile,"La fenêtre " & $hwndAccueil & " n'est pas active.")
			EndIf
	EndIf
	If WinWaitActive($hwndAccueil, "", 15) Then ; Attend jusqu'à 15 secondes pour que la fenêtre soit active
		; Instructions à exécuter si la fenêtre est active
		WriteToLog($LogFile," cmd ctrl+o")
		If ControlSend($hwndAccueil, "", "[CLASS:view.apptitle]", "^o") = 1 Then
			WriteToLog($LogFile," cmd ctrl+o réussi")
		Else
			WriteToLog($LogFile," échec de la commande cmd ctrl+o avec ControleSend(), essai avec la cmd Send()")
			Send("^o")
		EndIf
	Else
		Send("^w")
		WriteToLog($LogFile,"La fenetre Sage 100 Gestion commerciale Premium n'a pas été actif avant 15s, fermeture...")
		WinClose($hwndAccueil)
		WinClose("[CLASS:GeCoMaes]")
		EndScript(0)
	EndIf

	$hwndOuvrirGcm =  GetHandle("#32770")  
	if($hwndOuvrirGcm = False) Then
		;fermeture de la base sage
		Send("^w")
		WinClose($hwndAccueil)
		EndScript(0)
	EndIf

	WriteToLog($LogFile,"En attente de la fenetre: ouvrir le fichier commercial...")
	WinActivate($hwndOuvrirGcm)
	If WinWaitActive($hwndOuvrirGcm, "", 15) Then   		

		WriteToLog($LogFile,"la fenetre ouvrir le fichier commercial est actif!")
		; Sélectionner le champ de texte "Nom du fichier"
		ControlFocus($hwndOuvrirGcm, "", "Edit1")
		; Entrer le chemin et le nom du fichier
		ControlSetText($hwndOuvrirGcm, "", "Edit1", $fileDbPath)
		; Cliquer sur le bouton "Ouvrir"
		ControlClick($hwndOuvrirGcm, "", "Button1")
		WriteToLog($LogFile," cmd button1")
	Else
		WriteToLog($LogFile,"La fenetre Ouvrir le fichier commercial n'a pas été actif avant 20s, fermeture...")
		WinClose($hwndAccueil)
		;ferme sage et le script
		EndScript(0)
	EndIf

	ControlClick($hwndOuvrirGcm, "Ou&vrir", "Button1")
	WriteToLog($LogFile,"Sage 100 est connecté à la base de données " & $fileDbPath)
	Sleep(2000)
;***********************************************************************************************************************************************************
EndFunc


Func DecodeUrl($url)
    $url = StringRegExpReplace($url, '%20', ' ')
	$url = StringRegExpReplace($url, '%2520', ' ')
    $url = StringRegExpReplace($url, '%5C', '\\')
	$url = StringRegExpReplace($url, '%255C', '\\')
	$url = StringRegExpReplace($url, '%2F', '\\')
	$url = StringRegExpReplace($url, '%3A', ':')
	$url = StringRegExpReplace($url, '%253A', ':')
    Return $url
EndFunc


Func DeleteLockFile($FilePath)
	If FileExists($FilePath) Then
		FileDelete($FilePath)
	EndIf
EndFunc  

Func EndScript($codeDeSortie)
	DeleteLockFile($LockFile)
	Exit($codeDeSortie) ; Arrête le script
EndFunc  

Func GetHandle($CLASS)
    Local $hWnd = WinGetHandle("[CLASS:" & $CLASS & "]")
    If @error Then
        WriteToLog($logFile, "Une erreur est survenue lors de la récupération du handle de la fenêtre " & $CLASS)
		Return False
    EndIf
	WriteToLog($logFile, "Récupération du handle" & $hWnd & " de la fenêtre " & $CLASS & " réussie")
    Return $hWnd
EndFunc


Func WriteToLog($logFilePath, $message)

	Local $hFile = -1

    If Not FileExists($logFilePath) Then
		;créer le fichier
		$hFile = FileOpen($logFilePath, 2) ; 2 = mode écriture
    	FileWriteLine($hFile, @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " - " & "----- Création du fichier log -----")
	EndIf
    
	$hFile = FileOpen($logFilePath, 1)  ; 1 = mode lecture
    ; Formatte le message avec la date et l'heure
    Local $formattedMessage = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & " - " & $message

    ; Écrit le message formaté dans le fichier de log
    FileWriteLine($hFile, $formattedMessage)

    ; Ferme le fichier de log
    FileClose($hFile)
EndFunc



;continue l'éxé du script si la fenetre actuelle est détecté, sinon la function essaye de se deconnecter
; de la base si il detecte la fenetre préccédente, sinon fermeture de la fenetre
Func EtapeDeVerification($fenetreActuel)
    WriteToLog($LogFile, "Etape de vérification")
    WriteToLog($logFile, $fenetreActuel)
    If WinExists($fenetreActuel) Then
		WinActivate($fenetreActuel)
		WriteToLog($LogFile, $fenetreActuel & "existe")
        If WinWaitActive($fenetreActuel, 10) Then 
            WriteToLog($LogFile,"la fenetre "& $fenetreActuel & " est actif!")
            ; Si la fenêtre est active, sortir de la fonction et continuer l'éxécution du code normalement
            return
        Else 
            WriteToLog($LogFile,"La fenêtre "& $fenetreActuel & " existe mais n'a pas été détectée après 10 secondes!")
            WinClose($fenetreActuel)
			Send("^w")
            EndScript(0)
        EndIf    
    Else 
	 ; fermeture du programme car on ne peux pas faire la commande pour se déconencter
		WriteToLog($LogFile, "fenetre précédente introuvable, fermeture du programme")
        WriteToLog($LogFile,"La fenêtre "& $fenetreActuel & " n'existe pas!")
		WinClose($hwndAccueil)
        EndScript(0)
    EndIf
EndFunc
