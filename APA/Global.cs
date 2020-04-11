using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Net.Mail;

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

        public static string NotenUmrechnen(string klasse, string note)
        {
            if (klasse.StartsWith("G"))
            {
                if (note == null || note == "")
                {
                    return "";
                }
                return note.Split('.')[0];
            }
            if (note == "15.0")
            {
                return "1+";
            }
            if (note == "14.0")
            {
                return "1";
            }
            if (note == "13.0")
            {
                return "1-";
            }
            if (note == "12.0")
            {
                return "2+";
            }
            if (note == "11.0")
            {
                return "2";
            }
            if (note == "10.0")
            {
                return "2-";
            }
            if (note == "9.0")
            {
                return "3+";
            }
            if (note == "8.0")
            {
                return "3";
            }
            if (note == "7.0")
            {
                return "3-";
            }
            if (note == "6.0")
            {
                return "4+";
            }
            if (note == "5.0")
            {
                return "4";
            }
            if (note == "4.0")
            {
                return "4-";
            }
            if (note == "3.0")
            {
                return "5+";
            }
            if (note == "2.0")
            {
                return "5";
            }
            if (note == "1.0")
            {
                return "5-";
            }
            if (note == "81.0")
            {
                return "Attest";
            }
            if (note == "99.0")
            {
                return "k.N.";
            }
            if (note == "0.0")
            {
                return "6";
            }
            Console.WriteLine("Fehler! Note nicht definiert!");
            Console.ReadKey();
            return "";
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
                return @" APA | Published under the terms of GPLv3 | Stefan Bäumer 2020 | Version 20200412".PadRight(50, '=');
            }
        }

        public static List<string> AbschlussKlassen
        {
            get
            {
                //return new List<string>() { "HHO", "HBTO", "HBFGO", "BSO", "12" };
                return new List<string>() { "GE13", "GW13", "GT13" };
            }
        }

        public static DateTime APA = new DateTime(2020, 04, 21);

        public static string Ziel = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\APA-" + Global.APA.Year + Global.APA.Month + Global.APA.Day + ".xlsx";

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }
        
        internal static void MailSenden(List<Lehrer> klassenleitungen, string subject, string body, string dateiname, byte[] attach)
        {
            ExchangeService exchangeService = new ExchangeService()
            {
                UseDefaultCredentials = true,
                TraceEnabled = false,
                TraceFlags = TraceFlags.All,
                Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx")
            };
            EmailMessage message = new EmailMessage(exchangeService);

            foreach (var item in klassenleitungen)
            {
                message.ToRecipients.Add(item.Mail);
            }
                        
            message.Subject = subject;

            message.Body = body;
            message.Attachments.AddFileAttachment(dateiname, attach);
            
            //message.SendAndSaveCopy();
            message.Save(WellKnownFolderName.Drafts);
            Console.WriteLine("            ... per Mail gesendet.");
            Console.ReadKey();
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
        }
    }
}