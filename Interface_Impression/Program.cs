﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IniParser;
using IniParser.Model;
using Objets100cLib;
using Newtonsoft.Json;
using System.Data.SqlClient;

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

            //ouverture des bases de données
            Console.WriteLine("tentative de connexion a la base SAGE...");

            SageCommandeManager sage = new SageCommandeManager(paramBaseCompta, paramBaseCial);

            if (sage.isconnected)
            {
                //recuperer les bons de livraison
                listDoc = sage.GetBonLivraison();
            }

            /************Connnexion bdd pour requete sql***************/
            SqlManager sqlManager = new SqlManager(sqlServerName, sqlServerDb, sqlServerUser, sqlServerPwd);

            foreach (String docPiece in listDoc)
            {
                
                string req = "SELECT CM_Modele FROM F_COMPTETMODELE " +
               "WHERE CT_Num = (SELECT DO_Tiers FROM F_DOCENTETE WHERE DO_Piece = '" + docPiece + "' " +
               "AND DO_Date = (SELECT MAX(DO_Date) FROM F_DOCENTETE WHERE DO_Piece = '" + docPiece + "')" +
               " AND CM_Type = 3)";


                //récupere le modele de document par rapport au numéro du bon de livraison
                //Console.WriteLine("modele:" + sqlManager.ExecuteSqlQuery(req));
                listModele.Add(Uri.EscapeDataString(Uri.EscapeDataString(sqlManager.ExecuteSqlQuery(req))));
                Console.WriteLine($"Modele : {sqlManager.ExecuteSqlQuery(req)}");
            }
            sqlManager.CloseConnexion();
            // Convertir la liste en JSON
            jsonDoc = JsonConvert.SerializeObject(listDoc);
            jsonModel = JsonConvert.SerializeObject(listModele);
            Console.ReadLine();

            //foreach (Dictionary<string, string> d in listPieceModele)
            //{
            //    foreach (KeyValuePair<string, string> kvp in dict)
            //    {
            //        Console.WriteLine($"Clé : {kvp.Key}, Valeur : {kvp.Value}");
            //    }
            //}

            /**********************************************************/
            Console.ReadLine();
            Console.WriteLine("éxecution du script autoit");
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

            // Tuer le processus AutoIt si le processus C# s'est arrêté
            if (!process.HasExited)
            {
                process.Kill();
            }
            Console.ReadLine();
        }
    }
}
