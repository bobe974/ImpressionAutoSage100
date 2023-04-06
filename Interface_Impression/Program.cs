using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
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
            //créer le fichier log
            Logger logger = new Logger();
            List<string> listDoc = new List<string>();
            List<string> listModele = new List<string>();
            List<object> listCbMarq = new List<object>();
            string jsonDoc, jsonModel = null;

            //lecture du fichier INI pour charger les infos de connexion
            Console.WriteLine(inifilePath);
            logger.WriteToLog("*************************************************************" +
                "");
            logger.WriteToLog("Début du programme");
          
            if (File.Exists(inifilePath))
            {
                string[] lines = File.ReadAllLines(inifilePath);
            }
            else
            {
                logger.WriteToLog("Le fichier INI n'existe, création du fichier");
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

            /********************************************Main******************************************/

            List<AutoItImprim> listImprim = new List<AutoItImprim>();

            Console.WriteLine($" Instaciation des objets d'impressions");
            // Instanciation des objets AutoIyImprim pour chaque base de données
            AutoItImprim autoIyImprim1 = new AutoItImprim(inifilePath, new Logger());
            AutoItImprim autoIyImprim2 = new AutoItImprim(inifilePath, new Logger());
            AutoItImprim autoIyImprim3 = new AutoItImprim(inifilePath, new Logger());

            listImprim.Add(autoIyImprim1);
            listImprim.Add(autoIyImprim2);
            listImprim.Add(autoIyImprim3);

            Console.WriteLine("parcours de la table d'impression de chaque bdd");
            //vérifie les bases sur laquelle il y a des documents à imprimer
            foreach(AutoItImprim objAuto in listImprim)
            {
                if (objAuto.VerifierTableImpression())
                {
                    Console.WriteLine($" la table d'impression de la bdd {objAuto.getdbName()} a des éléments a imprimé");
                    int maxEssais = 10; // nombre maximal d'essais
                    int delaiEntreEssais = 500; // temps d'attente entre chaque essai en millisecondes
                    int tempsEcoule = 0; // temps écoulé en millisecondes
                    bool imprimeTermine = false;
                    int essaiCourant = 0;

                    // lancer le processus d'impression 
                    Console.WriteLine($"***********************lancement du processus d'impression sur la base {objAuto.getdbName()}***********************");
                    objAuto.ImprimProcess();
                   //attendre que l'exécution du script retourne true pour passer a la base suivante
                    while (!imprimeTermine && essaiCourant < maxEssais && tempsEcoule < 30000)
                    {
                        imprimeTermine = objAuto.etatImpressionTerminee;
                        if (!imprimeTermine)
                        {
                            Thread.Sleep(delaiEntreEssais); // attendre avant de réessayer
                            essaiCourant++;
                            tempsEcoule += delaiEntreEssais;
                        }
                       
                    }
                    if (imprimeTermine)
                    {
                        Console.WriteLine($"fin d'exécution sur la base {objAuto.getdbName()}");
                    }
                    else
                    {
                        Console.WriteLine($"l'exécution sur la base  {objAuto.getdbName()} ne c'est pas terminé, timeout");
                    }
                }
                else
                {
                    Console.WriteLine($" la table d'impression de la bdd {objAuto.getdbName()} n'as pas d'élement a imprimer");
                }
                           
            }

            Console.ReadLine();

        }
    }
}
