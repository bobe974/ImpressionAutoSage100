using Objets100cLib;
using System;
using System.Collections.Generic;

namespace Interface_Impression
{
    class SageCommandeManager
    {
        //base comptable 
        private BSCPTAApplication100c dbCompta;
        //base commerciale 
        private BSCIALApplication100c dbCommerce;

        private string dbname = null;
        public bool isconnected = false;

        public SageCommandeManager(ParamDb paramCompta, ParamDb paramCommercial)
        {
            //initialisation et connexion aux bases sage100
            this.dbCompta = new BSCPTAApplication100c();
            this.dbCommerce = new BSCIALApplication100c();
            //récupere le nom de la base pour faire des traitements personnalisés
            this.dbname = paramCommercial.getName();
            if (OpenDbComptable(dbCompta, paramCompta) && (OpenDbCommercial(dbCommerce, paramCommercial, dbCompta)))
            {
                isconnected = true;
            }
        }
        
        public List<String> GetBonLivraison()
        {
            List<string> listDoc = new List<string>();
            try {
                Console.WriteLine("les bons de livraisons");
                //IBODocument3 bonLivraison = dbCommerce.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteLivraison, "");
                IBICollection collDoc = null;
                collDoc = dbCommerce.FactoryDocumentVente.QueryTypeOrderPiece(DocumentType.DocumentTypeVenteLivraison);

                foreach (IBODocument3 doc in collDoc)
                {          
                    //recupere uniquement les bl avec le statut non imprimé
                    if(doc.DO_Imprim == false)
                    {
                        Console.WriteLine($"nom {doc.DO_Piece} statut {doc.DO_Imprim}");
                        listDoc.Add(doc.DO_Piece);                         
                    }          
                }
                return listDoc;
            }
            catch(Exception e) {
                Console.WriteLine(e);
                return null;
            }
            Console.ReadLine();
        }

 
        bool OpenDbComptable(BSCPTAApplication100c dbComptable, ParamDb paramCpta)
        {
            try
            {
                //ouverture de la base comptable 
                paramCpta.getName();
                dbComptable.Name = paramCpta.getDbPath();
                dbComptable.Loggable.UserName = paramCpta.getuser();
                dbComptable.Loggable.UserPwd = paramCpta.getpwd();

                dbComptable.Open();
                Console.WriteLine("succes connexion à " + paramCpta.getDbPath());
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("erreur pour l'ouverture de la base comptable: " + e);
                return false;
            }
        }

        bool OpenDbCommercial(BSCIALApplication100c dbCommercial, ParamDb paramCial, BSCPTAApplication100c bdcompta)
        {
            try
            {
                //ouverture de la base comptable 
                dbCommercial.Name = paramCial.getDbPath();
                dbCommercial.Loggable.UserName = paramCial.getuser();
                dbCommercial.Loggable.UserPwd = paramCial.getpwd();
                dbCommercial.CptaApplication = bdcompta;
                dbCommercial.Open();
                Console.WriteLine("succes connexion à " + paramCial.getDbPath());
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("erreur pour l'ouverture de la base commerciale: " + e);
                return false;
            }
        }

        public bool CloseDB()
        {
            try
            {
                this.dbCompta.Close();
                this.dbCommerce.Close();
                Console.WriteLine("base fermée");
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }
    }
}
