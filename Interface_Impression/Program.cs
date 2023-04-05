using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using IniParser;
using IniParser.Model;
using Newtonsoft.Json;

namespace Interface_Impression
{
    class Program
    {
        //chemin du fichier INI
        public static string inifilePath = Path.Combine(Directory.GetCurrentDirectory(), "config.ini");

        static void Main(string[] args)
        {
            string scriptName = @"C:\Users\Utilisateur\Desktop\projet impression\autoIT\Test.au3";
            string parameter1 = @"C:\Program Files (x86)\Sage\Gestion commerciale 100c\gecomaes.exe";
            string parameter2 = @"C:\Users\Utilisateur\Desktop\Projet 1\STOCKSERVICE\STOCKSERVICE.gcm";
            string autoItPath = "C:\\Program Files (x86)\\AutoIt3\\AutoIt3.exe";
           
            List<string> listDoc = new List<string>();
            List<string> listModele = new List<string>();
            List<object> listCbMarq = new List<object>();
            string jsonDoc, jsonModel = null;

            //lecture du fichier INI pour charger les infos de connexion
            Console.WriteLine(inifilePath);
            if (File.Exists(inifilePath))
            {
                string[] lines = File.ReadAllLines(inifilePath);
            }
            else
            {
                Console.WriteLine("Le fichier n'existe pas, création du document...");
                IniData ini = new IniData();
                ini.Sections.AddSection("DatabaseComptaSage");
                ini["DatabaseComptaSage"].AddKey("Path", "xxxx");
                ini["DatabaseComptaSage"].AddKey("User", "xxxx");
                ini["DatabaseComptaSage"].AddKey("Password", "xxxx");

                ini.Sections.AddSection("DatabaseCommercialeSage");
                ini["DatabaseCommercialeSage"].AddKey("Path", "xxxx");
                ini["DatabaseCommercialeSage"].AddKey("User", "xxxx");
                ini["DatabaseCommercialeSage"].AddKey("Password", "xxxx");

                ini.Sections.AddSection("DatabaseSqlServer");
                ini["DatabaseSqlServer"].AddKey("ServerName", "xxxx");
                ini["DatabaseSqlServer"].AddKey("Database", "xxxx");
                ini["DatabaseSqlServer"].AddKey("User", "xxxx");
                ini["DatabaseSqlServer"].AddKey("Password", "xxxx");

                //serveur distant ou sont stockké les modèles de fichier
                ini.Sections.AddSection("cheminServeur");
                ini["cheminServeur"].AddKey("Path", "xxxx");

                //création du fichier
                FileIniDataParser fileParser = new FileIniDataParser();
                fileParser.WriteFile(inifilePath, ini);
            }
           
            Console.ReadLine();
            //Création d'un objet parser pour lire le fichier INI
            var parser = new FileIniDataParser();

            // Lecture du fichier INI
            IniData data = parser.ReadFile(inifilePath);
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

            string cheminServeur = data["cheminServeur"]["Path"];

            Console.WriteLine($"INI:{dbComptaPath}, {dbComptaUser}{dbComptaPwd}");
            Console.ReadLine();

            //connection a la base
            // Paramètres pour se connecter aux bases
            ParamDb paramBaseCompta = new ParamDb(dbComptaPath, dbComptaUser, dbComptaPwd);
            ParamDb paramBaseCial = new ParamDb(dbCialPath, dbCialUser, dbCialPwd);

            /************Connnexion bdd pour requete sql***************/
            SqlManager sqlManager = new SqlManager(sqlServerName, sqlServerDb, sqlServerUser, sqlServerPwd);

            //regarder si il y a des documents a imprimé dans la table de donénes
            if (sqlManager.GetRowCount("sog_impression")!= 0)
            {
                //recupere le cbmarque de tout les documents
                List<object> ListcbMarq = sqlManager.ExecuteQueryToList("select DISTINCT cbMarq from sog_impression");

                //on recupére le numéros de piece de chaque bon de commande
                foreach (object cbMarq in ListcbMarq)
                {
                    listDoc.Add(sqlManager.ExecuteSqlQuery("select DO_Piece from F_DOCENTETE where cbMarq =" + cbMarq.ToString()).ToString());
                    listCbMarq.Add(cbMarq);
                }
            }

            Console.WriteLine("contenu de la liste listnumDOC:");

            Console.ReadLine();
            foreach(string s in listDoc)
            {
                Console.WriteLine(s);
            }
            Console.ReadLine();
            
            //sqlManager.CloseConnexion();

            // Convertir la liste en JSON
            jsonDoc = JsonConvert.SerializeObject(listDoc);
            jsonModel = JsonConvert.SerializeObject(listModele);
      
            Console.ReadLine();
            Console.WriteLine("Execution du script autoit");
            Console.WriteLine(paramBaseCial.getName());

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = autoItPath;
            //startInfo.Arguments = "\"" + scriptName + "\" \"" + parameter1 + "\" \"" + parameter2 + "\"";
            startInfo.Arguments = "\"" + scriptName + "\" \"" + parameter1 + "\" \"" + parameter2 + "\" \"" + jsonDoc + "\" \"" + jsonModel + "\" \"" + paramBaseCial.getName() + "\"";

            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start(); 

            // Attendre que le processus AutoIt se termine
            process.WaitForExit(); 

            int exitCode = process.ExitCode; 
            process.Kill();

            if (exitCode == 1)
                {
                Console.WriteLine("La tâche a été effectuée avec succès");
                Console.WriteLine("******** Vérification que les bons de livraison ont bien été imprimés *********");
                //vérifie que le bon de livraison est bien passé a l'état imprimé dans sage
                //
                int compteur = 0; 

                //parcours des bons de livraison qu'on devait imprimer
                foreach(object numPiece in listDoc)
                {
                    Console.WriteLine($"num piece:  {numPiece}");
                    //recupere l'état d'impression d'un document
                    int doImprim = Convert.ToInt32(sqlManager.ExecuteSqlQuery("SELECT \"DO_Imprim\" FROM F_DOCENTETE WHERE \"DO_Piece\" = '" + numPiece + "'"));

                    //string doImprim = sqlManager.ExecuteSqlQuery("select DO_Imprim from F_DOCENTETE WHERE DO_Piece =" + numPiece);
                    Console.WriteLine($"état d'impression du document {numPiece}: {doImprim}");
                    // si = 1 le document a été imprimé
                    if (doImprim.Equals(1))
                    {
                        Console.WriteLine($"{numPiece} à été imprimé, supression dans la table...");                                           
                        //Supprimer la table de bon de livraison
                        sqlManager.deleteRow("sog_impression", listCbMarq[compteur].ToString());
                        Console.WriteLine($"la ligne {listCbMarq[compteur].ToString()} a été supprimé");                      
                    }
                    compteur++;
                }               
            }
            else
            {
                Console.WriteLine("La tâche a échoué avec le code de sortie : " + exitCode);
            }

            Console.ReadLine();
            //fermeture de la base de données
            sqlManager.CloseConnexion();
        }
    }
}
