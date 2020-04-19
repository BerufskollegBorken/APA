using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace APA
{
    public class Excelzeilen : List<Excelzeile>
    {
        public Excelzeilen()
        {
        }

        internal void ToExchange(Lehrers lehrers)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013)
            {
                UseDefaultCredentials = true
            };

            service.TraceEnabled = false;
            service.TraceFlags = TraceFlags.All;
            service.Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx");

            Excelzeilen e = new Excelzeilen();

            foreach (var lehrer in lehrers)
            {
                foreach (var excelzeile in this)
                {
                    foreach (var v in excelzeile.IVerantwortlich)
                    {
                        if (v.Kürzel == lehrer.Kürzel)
                        {
                            v.Excelzeilen.Add(excelzeile);
                        }
                    }
                }
            }

            foreach (var lehrer in lehrers)
            {
                if (lehrer.Excelzeilen.Count > 0)
                {
                    lehrer.ToExchange(service);
                }
            }
        }

        internal string ToClipboard()
        {
            //foreach (var item in this)
            //{
            //    Global.Clipboard += item.ADatum + "\t" + RenderTag(item.ADatum) + "\t" + RenderVonBis(item.ADatum, item.CVonBis, item.FBeginn, item.GEnde, item.JKategorie) + "\t" + RenderBeschreibung(item.DBeschreibung) + "\t" + item.EJahrgang + "\t" + (item.FBeginn.Hour == 0 ? "" : " " + item.FBeginn.ToString("HH:mm") + " ") + "\t" + (item.GEnde.Hour == 0 ? "" : " " + item.GEnde.ToString("HH:mm") + " ") + "\t" + RenderRaum(raums, item.HRaum) + "\t" + RenderVerantwortliche(lehrers) + "\t" + item.JKategorie + "\t" + RenderHinweis(item.KHinweise) + "\t" + item.LGeschützt + "\t" + item.MAnmerkungen + Environment.NewLine;
            //}
            return Global.Clipboard;
        }
        private string RenderBeschreibung(string dBeschreibung)
        {
            // Landeswapen bei Zentralprüfungen
            if (dBeschreibung.StartsWith("<b>Zentral"))
            {
                return "<img src='https://www.berufskolleg-borken.de/wp-content/uploads/2020/04/nrw-wappen.png'> " + dBeschreibung;
            }
            return dBeschreibung;
        }

        private string RenderBody(string beschreibung, string hinweise)
        {
            return beschreibung + "</br>" + hinweise;
        }

        private string RenderSubject(string beschreibung)
        {
            return beschreibung.Replace("</br>", " - ").Replace("<b>", "").Replace("</b>", "").Trim();
        }

        private dynamic RenderHinweis(string kHinweise)
        {
            return kHinweise.Replace("Beweglicher Ferientag", "<b><a href='https://www.berufskolleg-borken.de/beweglicheFerientage/'>Beweglicher Ferientag</a></b>");
        }

        private int GetKalenderwoche(DateTime datum)
        {
            CultureInfo CUI = CultureInfo.CurrentCulture;
            return CUI.Calendar.GetWeekOfYear(datum, CUI.DateTimeFormat.CalendarWeekRule, CUI.DateTimeFormat.FirstDayOfWeek);
        }

        private dynamic RenderKategorie(List<string> jKategorie)
        {
            var x = "";

            foreach (var item in jKategorie)
            {
                x += item + ",";
            }
            return x.TrimEnd(',');
        }
    }
}

        