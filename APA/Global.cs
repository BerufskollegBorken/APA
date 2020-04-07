using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;

namespace APA
{
    public static class Global
    {
        public static string InputNotenCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv";
        public static string InputExportLessons = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\ExportLessons.csv";
        public static string InputStudentgroupStudentsCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\StudentgroupStudents.csv";

        public static string ConAtl = @"Dsn=Atlantis9;uid=DBA";

        internal static void IstInputNotenCsvVorhanden()
        {
            if (!File.Exists(Global.InputNotenCsv))
            {
                RenderInputAbwesenheitenCsv(Global.InputNotenCsv);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(Global.InputNotenCsv).Date != DateTime.Now.Date)
                {
                    RenderInputAbwesenheitenCsv(Global.InputNotenCsv);
                }
            }

        }
        private static void RenderInputAbwesenheitenCsv(string inputNotenCsv)
        {
            Console.WriteLine("Die Datei " + inputNotenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Zeitraum definieren (z.B. Ganzes Schuljahr)");
            Console.WriteLine(" 3. Auf CSV-Ausgabe klicken");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static string ConU = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";

        public static string AdminMail { get; internal set; }

        public static string AktSjAtl
        {
            get
            {
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                return sj.ToString() + "/" + (sj + 1 - 2000);
            }
        }

        public static string AktSjUnt
        {
            get
            {
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                return sj.ToString() + (sj + 1);
            }
        }

        public static string Titel {
            get
            {
                return @" APA | Published under the terms of GPLv3 | Stefan Bäumer 2020 | Version 20200406\n".PadRight(50, '=');
            }
        }

        public static List<string> AbschlussKlassen
        {
            get
            {
                return new List<string>() { "HHO", "HBTO", "HBFGO", "BSO", "12" };
            }
        }

        public static string Ziel = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\APA1.xlsx";

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }

        internal static void MailSenden(Klasse to, string subject, string body, List<string> fileNames)
        {
            ExchangeService exchangeService = new ExchangeService()
            {
                UseDefaultCredentials = true,
                TraceEnabled = false,
                TraceFlags = TraceFlags.All,
                Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx")
            };
            EmailMessage message = new EmailMessage(exchangeService);

            foreach (var item in to.Klassenleitungen)
            {
                if (item.Mail != null && item.Mail != "")
                {
                    //message.ToRecipients.Add("baeumer@posteo.de");
                    message.ToRecipients.Add(item.Mail);
                }                
            }
            
            message.BccRecipients.Add("stefan.baeumer@berufskolleg-borken.de");

            message.Subject = subject;

            message.Body = body;
            
            foreach (var datei in fileNames)
            {                
                message.Attachments.AddFileAttachment(datei);
            }
            
            //message.SendAndSaveCopy();
            message.Save(WellKnownFolderName.Drafts);
            Console.WriteLine("            " + subject + " ... per Mail gesendet.");
        }
    }
}