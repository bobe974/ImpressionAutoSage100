using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            Console.WriteLine("filename: " + fileName);
            return fileName;
        }
    }
}
