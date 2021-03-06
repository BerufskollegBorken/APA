﻿using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace APA
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);
            
            try
            {   
                Console.WriteLine(Global.Titel);

                Global.IstInputNotenCsvVorhanden();
                
                var prds = new Periodes();
                var fchs = new Fachs();
                var lehs = new Lehrers(prds);
                var klss = new Klasses(lehs, prds);
                var schuelers = new Schuelers(klss, lehs);
                Excelzeilen excelzeilen = new Excelzeilen();
                schuelers.Unterrichte();                  
                excelzeilen.AddRange(klss.Notenlisten(schuelers, lehs));
                excelzeilen.ToExchange(lehs);
                lehs.FehlendeUndDoppelteEinträge(schuelers);                
                System.Diagnostics.Process.Start(Global.Ziel);
                //Global.MailSenden(new Klasse(), new Lehrer(), "Liste alle Dokumente für den APA", "Siehe Anlage.", klss.Dokumente());
                //System.Windows.Forms.Clipboard.SetText(excelzeilen.ToClipboard());
                Console.WriteLine("Tabelle ZulassungskonferenzBC in Zwischenablage geschrieben.");
                Console.ReadKey();
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimbam! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

                
        private static bool DateiGöffnet(string inputAbwesenheitenCsv)
        {
            try
            {

            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains(" , da sie von einem anderen Prozess verwendet wir"))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
