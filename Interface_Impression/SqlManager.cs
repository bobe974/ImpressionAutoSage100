using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace Interface_Impression
{
    class SqlManager : IDisposable
    {
        SqlConnection connection;

        public SqlManager(string serverName, string database, string username, string pwd)
        {
            string connectionString = "Data Source=" + serverName + ";Initial Catalog=" + database + ";User ID=" + username + ";Password=" + pwd;
            //string connectionString = "Data Source=" + serverName + ";Initial Catalog=" + database + ";Integrated Security=SSPI";
            connection = new SqlConnection(connectionString);
            try
            {
                connection.Open();
                Console.WriteLine("La connexion à la base de données a été établie avec succès.");
            }
            catch (Exception e)
            {
                Console.WriteLine("La connexion à la base de données a échoué : " + e.Message);
                Dispose();
            }
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        /**
         * exécute une requete select et retourne le premier element 
         */
        public string ExecuteSqlQuery(String query)
        {
            String modele = null;

            if (connection != null && connection.State == ConnectionState.Open)
            {
                SqlCommand command = new SqlCommand(query, connection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        modele = (string)reader.GetValue(0);
                    }
                }
                return modele;
            }
            else
            {
                throw new Exception("La connexion n'est pas ouverte.");
            }
        }

        /**
         * exectute la requete et retourne les elements du premier champ dans une liste
         */
        public List<object> ExecuteQueryToList(String query)
        {
            List<object> list = new List<object>();
            //String modele = null;

            if (connection != null && connection.State == ConnectionState.Open)
            {
                SqlCommand command = new SqlCommand(query, connection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add((reader.GetValue(0)));
                    }
                                
                    return list;
                }
            }
            else
            {
                throw new Exception("La connexion n'est pas ouverte.");
            }
        }

        public int GetRowCount(string table)
        {
            int rowCount = 0;

            if (connection != null && connection.State == ConnectionState.Open)
            {
                string query = "SELECT COUNT(*) FROM "+ table;
                SqlCommand command = new SqlCommand(query, connection);

                rowCount = (int)command.ExecuteScalar();
                Console.WriteLine(rowCount);
            }
            else
            {
                throw new Exception("La connexion n'est pas ouverte.");
            }

            return rowCount;
        }



        public void CloseConnexion()
        {
            connection.Close();
        }
    }
}
