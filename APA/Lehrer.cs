using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;

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
        public Excelzeilen Excelzeilen { get; internal set; }

        public Lehrer(string anrede, string vorname, string nachname, string kürzel, string mail, string raum)
        {
            Excelzeilen = new Excelzeilen();
            Anrede = anrede;
            Nachname = nachname;
            Vorname = vorname;
            Raum = raum;
            Mail = mail;
            Kürzel = kürzel;
        }

        public Lehrer()
        {
            Excelzeilen = new Excelzeilen();
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

            message.Body = @"Guten Tag " + this.Vorname + " " + this.Nachname + "," +
                "<br>" +
                                "<br>" +
                "Sie erhalten diese Mail, weil für folgende Schülerinnen und Schüler bisher keine Vornoten eingetragen wurden:" +
                                "<br>" + 
                                "<br><table>";
            {
                foreach (var schueler in (from ss in schuelerOhneNoten select ss).OrderBy(x => x.Klasse.NameUntis).ThenBy(x => x.Nachname))
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
<br>Sofern Sie noch eintragen müssen, holen Sie das bis spätestens 13.04.20 um 24 Uhr im Digitalen Klassenbuch nach.<br><br>Mit kollegialem Gruß<br>Stefan Bäumer";

                //message.SendAndSaveCopy();
                message.Save(WellKnownFolderName.Drafts);
                Console.WriteLine("            " + message.Subject + " " + this.Kürzel + " ... per Mail gesendet.");
            }
        }

        internal void ToExchange(ExchangeService service)
        {
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, this.Mail);

            Appointments appointmentsIst = new Appointments(this.Mail, service);

            Appointments appointmentsSoll = new Appointments(Excelzeilen, service);

            appointmentsIst.DeleteAppointments(appointmentsSoll);

            appointmentsSoll.AddAppointments(appointmentsIst, this, service);
        }
    }
}