using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace APA
{
    public class Excelzeile
    {
        public DateTime ADatum { get; set; }
        public string BTag { get; set; }
        public List<DateTime> CVonBis { get; set; }
        public string DBeschreibung { get; set; }
        public string EJahrgang { get; set; }
        public DateTime FBeginn { get; set; }
        public DateTime GEnde { get; set; }
        public Raums HRaum { get; set; }
        public List<Lehrer> IVerantwortlich { get; set; }
        public List<string> JKategorie { get; set; }
        public string KHinweise { get; set; }
        public string LGeschützt { get; set; }
        public string Subject { get; set; }
        public MessageBody Body { get; set; }
    }
}