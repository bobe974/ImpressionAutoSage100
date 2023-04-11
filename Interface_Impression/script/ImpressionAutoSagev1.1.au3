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
; Version: v1.01
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

;**************************************************Vérification de l'état de sage 100******************************************************
; TODO Gestion des fenetres pour chaque base
WinActivate ( "Sage 100 Gestion commerciale Premium")
; ouvre sage 100 si il n'est pas lancé
If WinExists("Sage 100 Gestion commerciale Premium") Or WinExists("STOCKSERVICE - Sage 100 Gestion commerciale Premium") Then
	;fermeture de la base actuelle 
	ControlSend("Sage 100 Gestion commerciale Premium", "", "[CLASS:view.apptitle]", "^w")
	WriteToLog($LogFile,"une fenetre sage 100 est déja ouverte")
Else
	;MsgBox(0, "Attention", "Sage100 n'est pas lancé.")
    ; Lance Sage100
    Run($filePath)
	WriteToLog($LogFile,"Lancement de l'exe " & $filePath)
EndIf
;***********************************************************************************************************************************************************

;**************************************************Ouverture du fichier commercial .gcm dans sage100******************************************************
WriteToLog($LogFile," tentative d'ouverture du fichier commercial .gcm")
WinActivate ( "Sage 100 Gestion commerciale Premium")
If WinWaitActive("Sage 100 Gestion commerciale Premium", "", 15) Then ; Attend jusqu'à 20 secondes pour que la fenêtre soit active
    ; Instructions à exécuter si la fenêtre est active
	ControlSend("Sage 100 Gestion commerciale Premium", "", "[CLASS:view.apptitle]", "^o")
	WriteToLog($LogFile," cmd ctrl o")
Else
	;Send("^w")
	WinClose("Sage 100 Gestion commerciale Premium")
	DeleteLockFile($LockFile)
	WriteToLog($LogFile,"La fenetre Sage 100 Gestion commerciale Premium n'a pas été actif avant 20s, fermeture...")
    Exit(0) ; Arrête le script
EndIf

WriteToLog($LogFile,"wait fenetre ouvrir le fichier commercial...")
If WinWaitActive("Ouvrir le fichier commercial", "", 20) Then   
	WriteToLog($LogFile,"la fenetre ouvrir le fichier commercial est actif!")
	; Sélectionner le champ de texte "Nom du fichier"
ControlFocus("Ouvrir le fichier commercial", "", "Edit1")
; Entrer le chemin et le nom du fichier
ControlSetText("Ouvrir le fichier commercial", "", "Edit1", $fileDbPath)
; Cliquer sur le bouton "Ouvrir"
ControlClick("Ouvrir le fichier commercial", "", "Button1")
WriteToLog($LogFile," cmd button1")

Else
	;Send("^w")
	WinClose("Sage 100 Gestion commerciale Premium")
	DeleteLockFile($LockFile)
    Exit(0) ; Arrête le script
EndIf

ControlClick("Ouvrir le fichier commercial", "Ou&vrir", "Button1")
WriteToLog($LogFile,"Sage 100 est connecté a la base de données " & $fileDbPath)
Sleep(2000)

;***********************************************************************************************************************************************************

;**************************************************Gestion du message bloquant sur la fermeture fiscal******************************************************
WriteToLog($LogFile,"gestion du message de fermeture fiscal")
Sleep(2000)
WinActivate ( "Sage 100 Gestion commerciale Premium")
If WinWaitActive("Sage 100 Gestion commerciale Premium","",20) Then 
	WriteToLog($LogFile,"fenetre de fermeture fiscal actif")
	Sleep(5000)
	WinActivate ("Sage 100 Gestion commerciale Premium")
	WriteToLog($LogFile,"Appuie sur ESC pour fermer fenetre bloquante")
	Send("{ESC}")
EndIf

Sleep(2000)
WriteToLog($LogFile,"Appuie sur ESC")
Send("{ESC}")

;***********************************************************************************************************************************************************

;**************************************************Ouverture de la fenetre "Liste des documents de vente"******************************************************
WriteToLog($LogFile,"Tentative d'ouverture de la fenetre liste des documents de vente")
WinActivate ($dbName&" - Sage 100 Gestion commerciale Premium")
WriteToLog($LogFile,"wait fenetre "& $dbName&" - Sage 100 Gestion commerciale Premium...")
If WinWaitActive($dbName&" - Sage 100 Gestion commerciale Premium", "", 20) Then 
	WriteToLog($LogFile,"la fenetre "& $dbName &" - Sage 100 Gestion commerciale Premium est actif!")
	Send("!t")
	WriteToLog($LogFile,"cmd alt+t")
	; Sélectionner l'item Documents de ventes
	Send("v")
	WriteToLog($LogFile,"cmd alt+v")
	Send("{ENTER}")	
Else 
	;Send("^w")
	WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
	DeleteLockFile($LockFile)
	WriteToLog($LogFile,"La fenêtre Liste des documents de vente n'a pas été détectée après 20 secondes, fermeture...")
    Exit(0) ; Arrête le script
EndIf
;***********************************************************************************************************************************************************

;**************************************************PROCESS IMPRESSION***************************************************************************************
; Activer la fenêtre de l'instance correcte
WinActivate($dbName&" - Sage 100 Gestion commerciale Premium")
If WinWaitActive($dbName&" - Sage 100 Gestion commerciale Premium","",20) Then ; Attend jusqu'à 20 secondes pour que la fenêtre soit active

; Attendre que la liste déroulante soit affichée
Sleep(1000)
WriteToLog($LogFile,"Processus d'impression...")
;Boucle sur tous les bon de livraison (récupéré depuis le programme C#)
For $i = 0 To UBound($docArray) - 1

	WriteToLog($LogFile,"***************Impression du document n°"& 0 & ": " & $docArray[$i] &"***************")
; Cliquer sur le bouton "Actions"
	ControlClick($dbName&" - Sage 100 Gestion commerciale Premium", "Actions", "view.menubutton6")
	WriteToLog($LogFile,"cmd btn Action")
	Sleep(2000)
	ControlSend($dbName&" - Sage 100 Gestion commerciale Premium","Actions","[CLASS:view.menubutton]","{DOWN 5}")
	WriteToLog($LogFile,"cmd down x5")
	ControlSend($dbName&" - Sage 100 Gestion commerciale Premium","Actions","[CLASS:view.menubutton]","{ENTER}")

	;***************Fenetre pour saisir les numéros des bons de livraisons;***************
	WriteToLog($LogFile,"wait fenetre Impression Liste de documents... ")
	WinWait("Impression Liste de documents")
	WriteToLog($LogFile,"la fenetre Impression Liste de documents existe! ")
	
	; Cliquer sur le bouton document
	Local $hWnd = WinWait("[CLASS:view.dialogwnd]", "", 10)
	;selectionner "document" dans la fenetre d'impression dans la combobox "type état"
	ControlClick($hWnd, "Liste", "ComboBox1")
	Sleep(2000)
	ControlSend("Impression Liste de documents","Liste", "ComboBox1","{DOWN 2}")
	ControlSend("Impression Liste de documents","Liste","ComboBox1","{ENTER}")

	; selection de l'option Bon de livraison dans le champs "Document"
	ControlClick($hWnd, "Tous", "ComboBox2")	
	Sleep(2000)
	ControlSend("Impression Liste de documents","Document", "ComboBox2","{DOWN 4}")
	ControlSend("Impression Liste de documents","Document", "ComboBox2","{ENTER}")
	Sleep(2000)
	
		;***************saisie des numéros des bons de livraisons;***************
	ControlClick("Impression Liste de documents", "", "Edit3")
	ControlSend("Impression Liste de documents", "", "Edit3", $docArray[$i])
	WriteToLog($LogFile,"saisie du numéro de document num:" & $docArray[$i])
	Sleep(2000)
	ControlClick("Impression Liste de documents", "", "Edit4")	
	ControlSend("Impression Liste de documents", "", "Edit4", $docArray[$i])
	Sleep(2000)
	ControlClick("Impression Liste de documents", "OK", "Button1")
	Sleep(6000)

	;Dans la cas ou la fenetre de selection de modele s'ouvre
	If WinExists("Sélectionner le modèle", "") Then 
		WriteToLog($LogFile,"La fenetre de selection de modele existe")
		WriteToLog($LogFile,"Le client lié au BL n'as pas de modele d'impression, fermeture...")
		ControlSend("Sélectionner le modèle", "", "[CLASS:#32770]", "{ESC}")
		WriteToLog($LogFile,"cmd Annuler")
	Else
		;Boutton d'impression    WARNING le programme peut rester bloquer si la fenetre d'impression n'apparait pas!
		;ControlClick("Impression", "OK", "Button19")
		Sleep(5000)
		WriteToLog($LogFile,"impression de :" & $docArray[$i])
	EndIf
Next

;******************************************************************Fin impression*********************************************************************

WriteToLog($LogFile,"fin d'impression")

;ferme sage 100
Send("^w")
;WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
DeleteLockFile($LockFile)
Exit(1) ; ferme le programme et retourne 1 au code c# pour confirmé la fin d"exécution du script

Else
	;Send("^w")
	WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
	DeleteLockFile($LockFile)
	WriteToLog($LogFile,"La fenetre "&$dbName&" - Sage 100 Gestion commerciale Premium n'a pas été active apres 20s")
    Exit(0)
EndIf

;******************************************************************FIN SCRIPT*********************************************************************

;FONCTIONS...

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


Func WriteToLog($logFilePath, $message)

	Local $hFile = -1

    If Not FileExists($logFilePath) Then
		;créer le fichier
		$hFile = FileOpen($logFilePath, 2) ; 2 = mode écriture
    	FileWriteLine($hFile, @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " - " & "----- Création du fichier log -----")
	EndIf
    
	$hFile = FileOpen($logFilePath, 1)  ; 1 = mode lecture
    ; Formatte le message avec la date et l'heure
    Local $formattedMessage = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " - " & $message

    ; Écrit le message formaté dans le fichier de log
    FileWriteLine($hFile, $formattedMessage)

    ; Ferme le fichier de log
    FileClose($hFile)
EndFunc

