#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <Json.au3>

;===============================================================================
;
; Nom du fichier: ImpressionAutoSage100.au3
; Auteur: Baillif Etienne
; Date de création: 07/04/2023
; Version: v1.00
; Description: Script pour automatiser l'impression de documents de bon de livraison dans Sage100
;
;===============================================================================


; Récupère le premier argument de ligne de commande comme chemin de fichier
$filePath = $CmdLine[1]
$fileDbPath = $CmdLine[2]
$jsonList = $CmdLine[3]
$dbName = $CmdLine[4]
$docArray = Json_Decode($jsonList)

; Lance sage100
Run($filePath) 

;**************************************************Ouverture du fichier commercial .gcm dans sage100******************************************************
WinActivate ( "Sage 100 Gestion commerciale Premium")
If WinWaitActive("Sage 100 Gestion commerciale Premium", "", 15) Then ; Attend jusqu'à 20 secondes pour que la fenêtre soit active
    ; Instructions à exécuter si la fenêtre est active
	ControlSend("Sage 100 Gestion commerciale Premium", "", "[CLASS:view.apptitle]", "^o")
Else
	WinClose("Sage 100 Gestion commerciale Premium")
    Exit("La fenetre Sage 100 Gestion commerciale Premium n'a pas été actif avant 20s") ; Arrête le script
EndIf

If WinWaitActive("Ouvrir le fichier commercial", "", 20) Then   
	; Sélectionner le champ de texte "Nom du fichier"
ControlFocus("Ouvrir le fichier commercial", "", "Edit1")
; Entrer le chemin et le nom du fichier
ControlSetText("Ouvrir le fichier commercial", "", "Edit1", $fileDbPath)
; Cliquer sur le bouton "Ouvrir"
ControlClick("Ouvrir le fichier commercial", "", "Button1")

Else
	WinClose("Sage 100 Gestion commerciale Premium")
    Exit(0) ; Arrête le script
EndIf

If WinWaitActive("Ouvrir le fichier commercial") Then 
    ; Instructions à exécuter si la fenêtre est active
	ControlClick("Ouvrir le fichier commercial", "Ou&vrir", "Button1")
	Sleep(2000)
Else

	WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
    Exit ; Arrête le script
EndIf
;***********************************************************************************************************************************************************

;**************************************************Gestion du message bloquant sur la fermeture fiscal******************************************************
If WinWaitActive("Sage 100 Gestion commerciale Premium","",20) Then 
	Sleep(3000)
	WinActivate ("Sage 100 Gestion commerciale Premium")
	Send("{ESC}")
EndIf
Send("{ESC}")

;***********************************************************************************************************************************************************

;**************************************************Ouverture de la fenetre "Liste des documents de vente"******************************************************

WinActivate ($dbName&" - Sage 100 Gestion commerciale Premium")
If WinWaitActive($dbName&" - Sage 100 Gestion commerciale Premium", "", 20) Then 
	Send("!t")
	; Sélectionner l'item Documents de ventes
	Send("v")
	Send("{ENTER}")	
Else 
	WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
    Exit("La fenêtre Liste des documents de vente n'a pas été détectée après 20 secondes") ; Arrête le script
EndIf
;***********************************************************************************************************************************************************

;**************************************************PROCESS IMPRESSION***************************************************************************************
; Activer la fenêtre de l'instance correcte
WinActivate($dbName&" - Sage 100 Gestion commerciale Premium")
If WinWaitActive($dbName&" - Sage 100 Gestion commerciale Premium","",20) Then ; Attend jusqu'à 20 secondes pour que la fenêtre soit active

; Attendre que la liste déroulante soit affichée
Sleep(1000)

;Boucle sur tous les bon de livraison (récupéré depuis le programme C#)
For $i = 0 To UBound($docArray) - 1

; Cliquer sur le bouton "Actions"
	ControlClick($dbName&" - Sage 100 Gestion commerciale Premium", "Actions", "view.menubutton6")
	Sleep(2000)
	ControlSend($dbName&" - Sage 100 Gestion commerciale Premium","Actions","[CLASS:view.menubutton]","{DOWN 5}")
	ControlSend($dbName&" - Sage 100 Gestion commerciale Premium","Actions","[CLASS:view.menubutton]","{ENTER}")
	
	;***************Fenetre pour saisir les numéros des bons de livraisons;***************
	WinWait("Impression Liste de documents")

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
	Sleep(2000)
	ControlClick("Impression Liste de documents", "", "Edit4")	
	ControlSend("Impression Liste de documents", "", "Edit4", $docArray[$i])
	Sleep(2000)
	ControlClick("Impression Liste de documents", "OK", "Button1")
	Sleep(2000)

	;Boutton d'impression    WARNING le programme peut rester bloquer si la fenetre d'impression n'apparait pas!
	;ControlClick("Impression", "OK", "Button19")
	Sleep(5000)
	;******************************************************************Fin impression*********************************************************************
Next

;ferme sage 100
WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
Exit(1) ; ferme le programme et retourne 1 au code c# pour confirmé la fin d"exécution du script

Else
	WinClose($dbName&" - Sage 100 Gestion commerciale Premium")
    Exit("La fenetre "&$dbName&" - Sage 100 Gestion commerciale Premium n'a pas été active apres 20s")
EndIf


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
