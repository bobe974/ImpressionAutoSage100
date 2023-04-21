
using System;
using System.IO;

/*
 * Nom de la classe : ParamDb
 * Description : Structure pour la connexion a une base de données
 * Auteur : Etienne Baillif
 * Date de création : 21/04/2023
 */

namespace Interface_Impression
{
    public class ParamDb
    {
        private String dbPath, user, pwd;

        public ParamDb(String dbPath, String user, String pwd)
        {
            this.dbPath = dbPath;
            this.user = user;
            this.pwd = pwd;
        }

        public String getDbPath()
        { return this.dbPath; }

        public String getuser()
        { return this.user; }

        public String getpwd()
        { return this.pwd; }

        public String getName()
        {
            string fileName = Path.GetFileNameWithoutExtension(dbPath);
            return fileName;
        }
    }
}
