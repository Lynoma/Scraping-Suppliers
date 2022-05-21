using S22.Imap;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using WinSCP;

namespace RetrieveWestcoast
{
    class Program
    {
        static void Main(string[] args)
        {
            execImport();
            getZip();
            GetPCAFromFTP();
            uploadFile();
        }

        private static void getZip()
        {
            #region vars
            string login = "*";
            string mdp = "*";
            string directory = Directory.GetCurrentDirectory();
            #endregion
            try
            {
                using (ImapClient client = new ImapClient("imap.gmail.com", 993, login, mdp, AuthMethod.Login, true))
                {
                    Console.WriteLine("connecté");

                    //Créé le dossier si il n'existe pas

                    if (!Directory.Exists(directory + "/downloads"))
                    {
                        Directory.CreateDirectory(directory + "/downloads");
                    }

                    //Supprime le fichier si il existe

                    if (File.Exists(directory + "/downloads/17130080STOCK.csv"))
                    {
                        File.Delete(directory + "/downloads/17130080STOCK.csv");
                    }

                    var msg = client.Search(SearchCondition.From("************")); //adresse a rechercher
                    uint lastmessage = 0;
                    foreach (var message in msg)
                    {
                        if (Convert.ToInt32(message) > lastmessage)
                            lastmessage = Convert.ToUInt32(message);
                    }
                    Console.WriteLine(lastmessage);

                    var attachment = client.GetMessage(lastmessage).Attachments[0];

                    //Download
                    using (var fileStream = File.Create(directory + "/downloads/import.zip"))
                    {
                        attachment.ContentStream.Seek(0, SeekOrigin.Begin);
                        attachment.ContentStream.CopyTo(fileStream);
                    }
                    Console.WriteLine("Téléchargé");
                    
                    ZipFile.ExtractToDirectory(directory + "/downloads/import.zip", directory + "/downloads");
                    File.Delete(directory + "/downloads/import.zip");
                }

                Console.WriteLine("déconnecté");
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }
        }

        private static void GetPCAFromFTP()
        {
            string directory = Directory.GetCurrentDirectory();
            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Ftp,
                HostName = "195.154.227.34",
                UserName = "2ALLOPANAS",
                Password = "a3!ZS7ykYe",
            };

            using (Session session = new Session())
            {
                session.Open(sessionOptions);
                session.GetFiles("2ALLOPANAS_pca_cat.csv", directory + "\\downloads\\2ALLOPANAS_pca_cat.csv").Check();
            }
        }

        private static void uploadFile()
        {
            string directory = Directory.GetCurrentDirectory();
            SessionOptions sessionOptions = new SessionOptions 
            {
                Protocol = Protocol.Ftp,
                HostName = "*",
                UserName = "*",
                Password = "*",
            };

            using (Session session = new Session())
            {
                session.Open(sessionOptions);

                Console.WriteLine("Connecté au serveur cpanel");

                session.PutFiles(directory + "\\downloads\\17130080STOCK.csv", "allofiestaloc.com/17130080STOCK.csv").Check();
                session.PutFiles(directory + "\\downloads\\tarif_1foteam.csv", "allofiestaloc.com/tarif_1foteam.csv").Check();
                session.PutFiles(directory + "\\downloads\\2ALLOPANAS_pca_cat.csv", "allofiestaloc.com/2ALLOPANAS_pca_cat.csv").Check();
            }
        }

        private static void execImport()
        {
            string directory = Directory.GetCurrentDirectory();
            var exelaunch = Process.Start(directory + "\\downloads\\ImportProject.exe");
            exelaunch.WaitForExit();
            File.Delete(directory + "\\produits.xls");
            File.Delete(directory + "\\1FotradeSupposedToWork.xlsx");
            if (File.Exists(directory + "/downloads/tarif_1foteam.csv"))
            {
                File.Delete(directory + "/downloads/tarif_1foteam.csv");
            }
            File.Move(directory + "\\1FotradeSupposedToWork.csv", directory + "\\downloads\\tarif_1foteam.csv");
        }
    }
}
