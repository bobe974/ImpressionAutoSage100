using System;
using System.Collections.Generic;
using System.Diagnostics;
using IniParser;
using IniParser.Model;
using Interface_Impression;
using Newtonsoft.Json;

/*
 * Nom de la classe : AutoItImprim
 * Description : Classe pour lancer et  intéragir avec le script AutoIT et vérifier que les documents on été imprimés
 * Auteur : Etienne Baillif
 * Date de création : 24/03/2023
 */

public class AutoItImprim
{
    List<string> listDoc = new List<string>();
    List<object> listCbMarq = new List<object>();
    string jsonDoc = null;
    SqlManager sqlManager = null;
    Logger logger = null;
    Logger logNonImprim = null;
    Logger logImprim = null;

    public bool etatImpressionTerminee { get; set; } = false;

    string scriptPath = null;  //chemin du script AutoIt
    string sageExePath = null; //chemin du .exe du logiciel sage
    string gcmFilePath = null; //fichier .gcm qui fait le lien entre sqlserver et sage
    string autoItPath = null;  //chemin du .exe du logiciel AutoIt
    string tableImpression = null; //nom de la table qui contient les documents a imprimer
    string sqlServerDb = null; //nom de la base de données

    public AutoItImprim(string iniFilePath, Logger logger)
    {
        this.logger = logger;
        this.logNonImprim = new Logger("DocumentNonImprimé.txt");
        this.logImprim = new Logger("DocumentImprimé.txt");

        //Création d'un objet parser pour lire le fichier INI
        var parser = new FileIniDataParser();

        // Lecture du fichier INI
        IniData data = parser.ReadFile(iniFilePath);
        logger.WriteToLog($"lecture du fichier ini");
        Console.WriteLine($"lecture du fichier ini");
        // Récupération des informations de connexion à partir du fichier INI

        string sqlServerName = data["DatabaseSqlServer"]["ServerName"];
        sqlServerDb = data["DatabaseSqlServer"]["Database"];
        string sqlServerUser = data["DatabaseSqlServer"]["User"];
        string sqlServerPwd = data["DatabaseSqlServer"]["Password"];
        tableImpression = data["DatabaseSqlServer"]["tableImpression"];

        scriptPath = data["AutoIt"]["ScriptPath"];
        autoItPath = data["AutoIt"]["ExeAutoItPath"];
        sageExePath = data["AutoIt"]["ExeSagePath"];
        gcmFilePath = data["AutoIt"][".gcmFilePath"];

        //connection a la base

        /************Connnexion bdd pour requete sql***************/
        this.sqlManager = new SqlManager(sqlServerName, sqlServerDb, sqlServerUser, sqlServerPwd);
    }

    public void ImprimProcess()
    {
        etatImpressionTerminee = false;
        //Si il y a des documents a imprimé -> processus d'impression
        //exécution de la procédure stocké qui permet de nettoyer la table sog impression des BL vides

        if (sqlManager.ExecuteStoredProcedure("sp_SOGEST_SUP_BL"))
        {
            Console.WriteLine("Exécution de la procédure stockée sp_SOGEST_SUP_BL Réussi");
            logger.WriteToLog("Exécution de la procédure stockée sp_SOGEST_SUP_BL Réussi");
        }
        else
        {
            Console.WriteLine("Echec d'éxécution  de la procédure stockée sp_SOGEST_SUP_BL");
            logger.WriteToLog("Echec d'éxécution  de la procédure stockée sp_SOGEST_SUP_BL");
        }
        if (sqlManager.GetRowCount(tableImpression) != 0)
        {
            //recupere le cbmarque de tout les documents
            List<object> ListcbMarq = sqlManager.ExecuteQueryToList($"select DISTINCT cbMarq from {tableImpression}");

            //on recupére le numéros de piece de chaque bon de commande
            foreach (object cbMarq in ListcbMarq)
            {
                string req = "select DO_Piece from F_DOCENTETE where cbMarq =" + cbMarq.ToString();

                try
                {
                    string doPiece = null;
                    //cas ou la requete retourne le doPiece 
                    if (!sqlManager.ExecuteSqlQuery(req).Equals(null))
                    {
                        doPiece = sqlManager.ExecuteSqlQuery(req).ToString();
                        listDoc.Add(doPiece);
                        listCbMarq.Add(cbMarq);
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine($"Erreur lors de la requete sql {req} : {e}");
                    logger.WriteToLog($"Erreur lors de la requete sql {req} : {e}");
                    //supprime de la table l'occurence dont aucun numéro de piece n'a été" trouvé depuis le champs cbMarq
                    sqlManager.deleteRow(tableImpression, cbMarq.ToString());
                    logNonImprim.WriteToLog(req);

                }
            }
            //afficher les bl à imprimer
            Console.WriteLine("BL à imprimer:");
            foreach (string s in listDoc)
            {
                Console.WriteLine(s);
            }

            if (!listDoc.Count.Equals(0))
            {
                // Convertir la liste en JSON, nécéssaire pour envoyer les données au script autoit...
                jsonDoc = JsonConvert.SerializeObject(listDoc);

                Console.WriteLine("//////////////////////Execution du script autoit...//////////////////////");
                logger.WriteToLog("//////////////////////Execution du script autoit...//////////////////////");

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = autoItPath;

                // le sageExePath est le premier element récupéré dans le script autoit...
                startInfo.Arguments = "\"" + scriptPath + "\" \"" + sageExePath + "\" \"" + gcmFilePath + "\" \"" + jsonDoc + "\" \"" + sqlServerDb + "\"";

                Process process = new Process();
                process.StartInfo = startInfo;
                process.Start();

                //attendre que le processus AutoIt se termine
                process.WaitForExit();

                int exitCode = process.ExitCode;
                //process.Kill();

                if (exitCode == 1)
                {
                    logger.WriteToLog("//////////////////////L'exécution du script autoIT s'est correctement terminé//////////////////////");
                    Console.WriteLine("//////////////////////L'exécution du script autoIT s'est correctement terminé//////////////////////");
                }
                else
                {
                    Console.WriteLine("//////////////////////La script autoit a échoué avec le code de sortie : : " + exitCode + "//////////////////////");
                    logger.WriteToLog("//////////////////////La script autoit a échoué avec le code de sortie : : " + exitCode + "//////////////////////");
                }

                // on vérifie si le script imprimé des documents, si oui, on supprime ces documents de la table d'impression
                VérificationDesDocummentImprimés();
            }
            else //cas ou la table d'impression est vide
            {
                logger.WriteToLog("aucun document n'a été trouvé dans la table d'impression");
                Console.WriteLine("aucun document n'a été trouvé dans la table d'impression");
            }

            //fermeture de la base de données
            sqlManager.CloseConnexion();
            logger.WriteToLog("fermeture de la base de données");
            //notifier que l'éxecution de la méthode est terminé
            etatImpressionTerminee = true;
        }

    }
    /**
     *  vérifie dans que la table qui contient les documents à imprimer n'est pas vide
     */
    public bool VerifierTableImpression() => sqlManager.GetRowCount(tableImpression) != 0;

    public void closeDB()
    {
        sqlManager.CloseConnexion();
    }

    public string getdbName()
    {
        return this.sqlServerDb;
    }

    /**
     * 1 vérifie que les documents on bien été imprimés et si c'est le cas on les supprime de la table d'impression
     * 2 ajoute dans un fichier les documents non imprimés
     */
    public bool VérificationDesDocummentImprimés()
    {
        Console.WriteLine("              Vérification que les BL ont bien été imprimés");
        //vérifie que le bon de livraison est bien passé a l'état imprimé dans sage
        bool isPrinted = true;
        int compteur = 0;
        int doImprim = 0;
        //parcours des bons de livraison qu'on devait imprimer
        foreach (object numPiece in listDoc)
        {
            Console.WriteLine($"num piece:  {numPiece}");
            //recupere l'état d'impression d'un document
            doImprim = 0;

            doImprim = Convert.ToInt32(sqlManager.ExecuteSqlQuery("SELECT \"DO_Imprim\" FROM F_DOCENTETE WHERE \"DO_Piece\" = '" + numPiece + "'"));

            Console.WriteLine($"état d'impression du document {numPiece}: {doImprim}");

            // si = 1 le document a été imprimé
                if (doImprim.Equals(1))
                {
                Console.WriteLine($"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!{numPiece} à été imprimé!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ");
                logImprim.WriteToLog(numPiece.ToString() + " " + sqlServerDb);  
                Console.WriteLine($"supression dans la table sog_Impression...");

                //Supprimer la table de bon de livraison
                sqlManager.deleteRow(tableImpression, listCbMarq[compteur].ToString());
                Console.WriteLine($"la ligne {listCbMarq[compteur].ToString()} a été supprimé de la table");
                logger.WriteToLog($"la ligne {listCbMarq[compteur].ToString()} a été supprimé de la table");

                /***************************MODIF**********************************/
                //change la valeur du champ libre impression pour éviter la réipression de documents lors d'une modification
                string query = "UPDATE F_DOCENTETE SET impression = 1  WHERE DO_Piece = '" + numPiece + "'";
                if (sqlManager.ExeSqlQuery(query))
                {
                    Console.WriteLine(@"le statut d'impression de la pièce {numPiece} a été changé à 1 (imprimé)");
                    logger.WriteToLog(@"le statut d'impression de la pièce {numPiece} a été changé à 1 (imprimé)");
                }
                else
                {
                    logger.WriteToLog(@"échec de changement du statut d'impression");
                }
                /***************************MODIF**********************************/
            }
            else
            {
                logNonImprim.WriteToLog(numPiece.ToString() + " " + sqlServerDb);
                logger.WriteToLog($"le bon de livraison DO_Piece: {numPiece}  et CbMarq: { listCbMarq[compteur].ToString()} n'a pas été imprimé");
                Console.WriteLine($"le bon de livraison DO_Piece: {numPiece}  et CbMarq: { listCbMarq[compteur].ToString()} n'a pas été imprimé");

                isPrinted = false;
            }
            compteur++;
        }
        return isPrinted;
    }
}
