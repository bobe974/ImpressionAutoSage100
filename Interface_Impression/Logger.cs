using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;


namespace Interface_Impression
{
    public class Logger
    {
       
        string logFilePath = "log.txt";

        public Logger()
        {
            // Vérifier si le fichier de log existe déjà
            if (!File.Exists(logFilePath))
            {
                // Créer un nouveau fichier de log
                using (TextWriterTraceListener listener = new TextWriterTraceListener(logFilePath))
                {
                    Trace.Listeners.Add(listener);
                    WriteToLog("Création du log");
                }
            }
        }

        public Logger(String filePath)
        {
            this.logFilePath = filePath;

            if (!File.Exists(logFilePath))
            {
                // Créer un nouveau fichier de log
                using (TextWriterTraceListener listener = new TextWriterTraceListener(logFilePath))
                {
                    Trace.Listeners.Add(listener);
                    WriteToLog("Création du log");
                }
            }
        }

        private string LogStringFormat(string msg)
        {
            DateTime now = DateTime.Now;
            string date = now.ToString("yyyy-MM-dd HH:mm:ss.fff");

            string format = date + " - " + msg;
            return format;
        }

        public void WriteToLog(string message)
        {
            // Ouvrir le fichier de log en mode append
            using (TextWriterTraceListener listener = new TextWriterTraceListener(logFilePath))
            {
                // Écrire dans le log
                listener.WriteLine(LogStringFormat(message));
                // Fermer le listener
                listener.Flush();
                listener.Close();
            }
        }

        public void WriteToFile(string content)
        {
            // Vérification des arguments
            if (string.IsNullOrEmpty(logFilePath))
                throw new ArgumentException("Le chemin de destination ne peut pas être null ou vide.");
            if (string.IsNullOrEmpty(content))
                throw new ArgumentException("Le contenu ne peut pas être null ou vide.");

            // Écriture du contenu dans le fichier texte
            using (StreamWriter writer = new StreamWriter(logFilePath))
            {
                writer.Write(content.Replace("\n", Environment.NewLine));
            }
        }

    }
}
