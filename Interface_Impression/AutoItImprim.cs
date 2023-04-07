using System;
using System.Collections.Generic;
using System.Diagnostics;
using IniParser;
using IniParser.Model;
using Interface_Impression;
using Newtonsoft.Json;

public class AutoItImprim
{
    List<string> listDoc = new List<string>();
    List<object> listCbMarq = new List<object>();
    string jsonDoc = null;
    SqlManager sqlManager = null;
    ParamDb paramBaseCompta = null;
    ParamDb paramBaseCial = null;
    Logger logger = null;
    public bool etatImpressionTerminee { get; set; } = false;

    string scriptPath = null;  //chemin du script AutoIt
    string sageExePath = null; //chemin du .exe du logiciel sage
    string gcmFilePath = null; //fichier .gcm qui fait le lien entre sqlserver et sage
    string autoItPath = null;  //chemin du .exe du logiciel AutoIt
    string tableImpression = null; //nom de la table qui contient les documents a imprimer

    public AutoItImprim(string iniFilePath, Logger logger)
    {

        this.logger = logger;
        //Création d'un objet parser pour lire le fichier INI
        var parser = new FileIniDataParser();

        // Lecture du fichier INI
        IniData data = parser.ReadFile(iniFilePath);
        logger.WriteToLog("lecture du fichier ini");
        Console.WriteLine("lecture du fichier ini...");
        // Récupération des informations de connexion à partir du fichier INI
        string dbComptaPath = data["DatabaseComptaSage"]["Path"];
        string dbComptaUser = data["DatabaseComptaSage"]["User"];
        string dbComptaPwd = data["DatabaseComptaSage"]["Password"];

        string dbCialPath = data["DatabaseCommercialeSage"]["Path"];
        string dbCialUser = data["DatabaseCommercialeSage"]["User"];
        string dbCialPwd = data["DatabaseCommercialeSage"]["Password"];

        string sqlServerName = data["DatabaseSqlServer"]["ServerName"];
        string sqlServerDb = data["DatabaseSqlServer"]["Database"];
        string sqlServerUser = data["DatabaseSqlServer"]["User"];
        string sqlServerPwd = data["DatabaseSqlServer"]["Password"];
        tableImpression = data["DatabaseSqlServer"]["tableImpression"];

        scriptPath = data["AutoIt"]["ScriptPath"];
        autoItPath = data["AutoIt"]["ExeAutoItPath"];
        sageExePath = data["AutoIt"]["ExeSagePath"];
        gcmFilePath = data["AutoIt"][".gcmFilePath"];
       

        Console.WriteLine($"INI:{dbComptaPath}, {dbComptaUser}{dbComptaPwd}");

        //connection a la base
        // Paramètres pour se connecter aux bases
        paramBaseCompta = new ParamDb(dbComptaPath, dbComptaUser, dbComptaPwd);
        paramBaseCial = new ParamDb(dbCialPath, dbCialUser, dbCialPwd);

        /************Connnexion bdd pour requete sql***************/
        this.sqlManager = new SqlManager(sqlServerName, sqlServerDb, sqlServerUser, sqlServerPwd);
    }

    public void ImprimProcess()
    {
        etatImpressionTerminee = false;
        //Si il y a des documents a imprimé -> processus d'impression
        if (sqlManager.GetRowCount(tableImpression) != 0)

        {
            //recupere le cbmarque de tout les documents
            List<object> ListcbMarq = sqlManager.ExecuteQueryToList($"select DISTINCT cbMarq from {tableImpression}");

            //on recupére le numéros de piece de chaque bon de commande
            foreach (object cbMarq in ListcbMarq)
            {
                Console.WriteLine(cbMarq);
                listDoc.Add(sqlManager.ExecuteSqlQuery("select DO_Piece from F_DOCENTETE where cbMarq =" + cbMarq.ToString()).ToString());
                listCbMarq.Add(cbMarq);
            }

            Console.WriteLine("contenu de la liste listnumDOC:");

            foreach (string s in listDoc)
            {
                Console.WriteLine(s);
            }

            // Convertir la liste en JSON
            jsonDoc = JsonConvert.SerializeObject(listDoc);

            Console.WriteLine("Execution du script autoit");
            logger.WriteToLog("Execution du script autoit");
            Console.WriteLine(paramBaseCial.getName());

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = autoItPath;

            // le sageExePath est le premier element récupéré dans le script autoit...
            startInfo.Arguments = "\"" + scriptPath + "\" \"" + sageExePath + "\" \"" + gcmFilePath + "\" \"" + jsonDoc + "\" \"" + paramBaseCial.getName() + "\"";

            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();

            // Attendre que le processus AutoIt se termine
            process.WaitForExit();

            int exitCode = process.ExitCode;
            //process.Kill();

            if (exitCode == 1)
            {
                logger.WriteToLog("L'exécution du script autoIT s'est correctement terminé");
                Console.WriteLine("La tâche a été effectuée avec succès");
            }
            else
            {
                Console.WriteLine("La tâche a échoué avec le code de sortie : " + exitCode);
                logger.WriteToLog("La script autoit a échoué avec le code de sortie : " + exitCode);
            }

            // on vérifie si le script imprimé des documents, meme en cas d'erreurs et si oui, on supprime ces documents de la table d'impression
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
        //notifié que la méthode est terminé
        etatImpressionTerminee = true;
    }
    public bool VerifierTableImpression() => sqlManager.GetRowCount(tableImpression) != 0;

    public void closeDB()
    {
        sqlManager.CloseConnexion();
    }

    public string getdbName()
    {
        return this.paramBaseCial.getName();
    }

    /**
     * vérifie que les documents on bien été imprimés et si c'es le cas on les supprime de la table d'impression
     */
    public bool VérificationDesDocummentImprimés()
    {
        Console.WriteLine("******** Vérification que les bons de livraison ont bien été imprimés *********");
        //vérifie que le bon de livraison est bien passé a l'état imprimé dans sage
        bool isPrinted = true;
        int compteur = 0;

        //parcours des bons de livraison qu'on devait imprimer
        foreach (object numPiece in listDoc)
        {
            Console.WriteLine($"num piece:  {numPiece}");
            //recupere l'état d'impression d'un document
            int doImprim = Convert.ToInt32(sqlManager.ExecuteSqlQuery("SELECT \"DO_Imprim\" FROM F_DOCENTETE WHERE \"DO_Piece\" = '" + numPiece + "'"));

            Console.WriteLine($"état d'impression du document {numPiece}: {doImprim}");
            // si = 1 le document a été imprimé
            if (doImprim.Equals(1))
            {
                Console.WriteLine($"{numPiece} à été imprimé, supression dans la table...");
                //Supprimer la table de bon de livraison
                sqlManager.deleteRow(tableImpression, listCbMarq[compteur].ToString());
                Console.WriteLine($"la ligne {listCbMarq[compteur].ToString()} a été supprimé de la table");
                logger.WriteToLog($"la ligne {listCbMarq[compteur].ToString()} a été supprimé de la table");
            }
            else
            {
                logger.WriteToLog($"le bon de livraison DO_Piece: {numPiece}  et CbMarq: { listCbMarq[compteur].ToString()} n'a pas été imprimé");
                Console.WriteLine($"le bon de livraison DO_Piece: {numPiece}  et CbMarq: { listCbMarq[compteur].ToString()} n'a pas été imprimé");
                isPrinted = false;
            }
            compteur++;
        }
        return isPrinted;
    }
}
