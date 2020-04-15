using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

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
    }
}