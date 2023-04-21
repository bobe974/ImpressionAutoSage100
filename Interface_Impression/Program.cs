using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

/*
 * Nom de la classe : Program
 * Description : Point d'entrée du programme -> 
 * Lecture des fichiers INI et vérifie toutes les x secondes si il y a des documents à imprimer.
 * Lance le processus d'impression sur chaque base de données trouvé dans les fichiers INI
 * Auteur : Etienne Baillif
 * Date de création : DateDeCreation
 */
namespace Interface_Impression
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(Directory.GetCurrentDirectory());

            // Boucle infinie qui répète le processus toutes les 20 secondes
            while (true)
            {
                //créer le fichier log
                Logger logger = new Logger();
                //lecture du fichier INI pour charger les infos de connexion
                logger.WriteToLog("");
                logger.WriteToLog("*****************************************Début de la vérification**************************************************");
                Console.WriteLine("");
                Console.WriteLine("*****************************************Début de la vérification**************************************************");

                //lecture de tout les fichiers ini dans le repertoire actuelle
                string directory = Directory.GetCurrentDirectory();
                List<AutoItImprim> listImprim = new List<AutoItImprim>();

                int i = 0;

                foreach (string filePath in Directory.GetFiles(directory, "*.ini"))
                {
                    if (File.Exists(filePath))
                    {
                        try {
                            Console.WriteLine("");
                            Console.WriteLine($"Traitement du fichier ini numero: {i + 1}");
                            logger.WriteToLog("");
                            logger.WriteToLog($"Traitement du fichier ini numero: {i + 1}");
                            listImprim.Add(new AutoItImprim(filePath, logger));
                        }
                        catch(Exception e)
                        {
                            Console.WriteLine(e);
                        }         
                    }
                    i++;
                }

                if (listImprim.Count != 0)
                {

                    /********************************************Main******************************************/

                    //vérifie les bases sur laquelle il y a des documents à imprimer
                    foreach (AutoItImprim objAuto in listImprim)
                    {
                        if (objAuto.VerifierTableImpression())
                        {
                            Console.WriteLine($" la table d'impression de la bdd {objAuto.getdbName()} contient des éléments à imprimer");
                            int maxEssais = 10; // nombre maximal d'essais
                            int delaiEntreEssais = 500; // temps d'attente entre chaque essai en millisecondes
                            int tempsEcoule = 0; // temps écoulé en millisecondes
                            bool imprimeTermine = false;
                            int essaiCourant = 0; 

                            // lancer le processus d'impression 
                            Console.WriteLine("");
                            Console.WriteLine($"***********************lancement du processus d'impression sur la base {objAuto.getdbName()}***********************");
                            objAuto.ImprimProcess();

                            //attendre que l'exécution du script retourne true pour passer a la base suivante
                            while (!imprimeTermine)                                                                 
                            {
                                imprimeTermine = objAuto.etatImpressionTerminee;
                                if (!imprimeTermine)
                                {
                                    Thread.Sleep(delaiEntreEssais);
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

                            Console.WriteLine("");
                            Console.WriteLine($"***********************Fin du processus d'impression sur {objAuto.getdbName()}***********************");
                        }
                        else
                        {
                            Console.WriteLine($" la table d'impression de la bdd {objAuto.getdbName()} n'as pas d'élement a imprimé");
                            logger.WriteToLog($" la table d'impression de la bdd {objAuto.getdbName()} n'as pas d'élement a imprimé");
                        }
                        
                    }

                }
                else
                {
                    Console.WriteLine("Aucun fichier ini trouvé");
                    logger.WriteToLog("Aucun fichier ini trouvé");
                }

                // Attendre 20 secondes
                logger.WriteToLog("");
                logger.WriteToLog("*****************************************Prochaine vérification dans 20s*******************************************");
                Console.WriteLine("");
                Console.WriteLine("*****************************************Prochaine vérification dans 20s*******************************************");
                Thread.Sleep(20000);
            }
              
        }
    }
}
