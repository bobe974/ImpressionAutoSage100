using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Interface_Impression
{
    class Program
    {

        static void Main(string[] args)
        {
            //créer le fichier log
            Logger logger = new Logger();
            //lecture du fichier INI pour charger les infos de connexion
            logger.WriteToLog("*************************************************************");
            logger.WriteToLog("Début du programme");

            //lecture de tout les fichiers ini dans le repertoire actuelle
            string directory = Directory.GetCurrentDirectory();
            List<AutoItImprim> listImprim = new List<AutoItImprim>();
            Console.WriteLine($" Instaciation des objets d'impressions");
            int i = 0;

            foreach (string filePath in Directory.GetFiles(directory, "*.ini"))
            {
                if (File.Exists(filePath))
                {
                    listImprim.Add(new AutoItImprim(filePath, logger));
                }
                i++;
            }

            if (listImprim.Count != 0)
            {

                Console.WriteLine($"Traitement du fichier ini: {i}");
                logger.WriteToLog($"Traitement du fichier ini: {i}");

                Console.ReadLine();

                /********************************************Main******************************************/

                Console.WriteLine("parcours de la table d'impression de chaque bdd");
                //vérifie les bases sur laquelle il y a des documents à imprimer
                foreach (AutoItImprim objAuto in listImprim)
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
                            Console.WriteLine($"début de boucle nbessai: {essaiCourant}");
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
                            logger.WriteToLog($"fin d'exécution sur la base {objAuto.getdbName()}");
                        }
                        else
                        {
                            Console.WriteLine($"l'exécution sur la base  {objAuto.getdbName()} ne c'est pas terminé, timeout");
                            logger.WriteToLog($"l'exécution sur la base  {objAuto.getdbName()} ne c'est pas terminé, timeout");
                        }
                    }
                    else
                    {
                        Console.WriteLine($" la table d'impression de la bdd {objAuto.getdbName()} n'as pas d'élement a imprimé");
                        logger.WriteToLog($" la table d'impression de la bdd {objAuto.getdbName()} n'as pas d'élement a imprimé");
                    }

                }
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Aucun fichier ini trouvé");
                logger.WriteToLog("Aucun fichier ini trouvé");
            }

        }
    }
}
