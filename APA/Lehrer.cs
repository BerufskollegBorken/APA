using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

namespace APA
{
    public class Lehrer
    {
        public int IdUntis { get; internal set; }
        public string Kürzel { get; internal set; }
        public string Mail { get; internal set; }
        public string Nachname { get; internal set; }
        public string Vorname { get; internal set; }
        public string Anrede { get; internal set; }
        public string Titel { get; internal set; }
        public string Raum { get; internal set; }
        public string Funktion { get; internal set; }
        public string Dienstgrad { get; internal set; }

        public Lehrer(string anrede, string vorname, string nachname, string kürzel, string mail, string raum)
        {
            Anrede = anrede;
            Nachname = nachname;
            Vorname = vorname;
            Raum = raum;
            Mail = mail;
            Kürzel = kürzel;
        }

        public Lehrer()
        {
        }

        internal void Mailen(List<Schueler> schuelerOhneNoten, List<Schueler> schuelerMitDoppelterNote)
        {
            ExchangeService exchangeService = new ExchangeService()
            {
                UseDefaultCredentials = true,
                TraceEnabled = false,
                TraceFlags = TraceFlags.All,
                Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx")
            };
            EmailMessage message = new EmailMessage(exchangeService);

            
            message.ToRecipients.Add(this.Mail);
            
            message.BccRecipients.Add("stefan.baeumer@berufskolleg-borken.de");

            message.Subject = "Fehlende Vornoten";

            message.Body = @"Guten Tag " + this.Vorname + " " + this.Nachname +"," +
                "<br>" +
                "Sie erhalten diese Mail, weil für folgende Schülerinnen und Schüler bisher keine Vornoten eingetragen wurden:" +
                "<br><table>";

            foreach (var schueler in schuelerOhneNoten)
            {
                foreach (var fach in schueler.Fächer)
                {
                    if (fach.Lehrerkürzel == this.Kürzel)
                    {
                        if (fach.Note == null || fach.Note == "")
                        {
                            message.Body += "<tr><td>" + schueler.Vorname + " " + schueler.Nachname.Substring(0, 1) + "</t><td>" + schueler.Klasse.NameUntis + "</td><td>" + fach.KürzelUntis + "</td></tr>";
                        }
                    }                    
                }                
            }

            message.Body += @"</table>
<br>Bitte holen Sie die Eintragungen bis spätestens " + DateTime.Now.AddHours(24).ToShortDateString() + "um 24 Uhr im Digitalen Klassenbuch nach.<br><br>Mit kollegialem Gruß<br>Stefan Bäumer";

            //message.SendAndSaveCopy();
            message.Save(WellKnownFolderName.Drafts);
            Console.WriteLine("            " + message.Subject + " " + this.Kürzel + " ... per Mail gesendet.");
        }
    }
}