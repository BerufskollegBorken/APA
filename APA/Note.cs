using System;
using System.Globalization;

namespace APA
{
    public class Note
    {
        public int StudentId { get; internal set; }
        public string Prüfungsart { get; private set; }
        public string LehrerKürzel { get; private set; }
        public string PrüfungsartNote { get; private set; }
        public string Klasse { get; internal set; }
        public DateTime Datum { get; internal set; }
        public string Fach { get; private set; }
        public object Lernbereich { get; internal set; }

        public Note(string line)
        {
            var x = line.Split('\t');
            Datum = GetDatum(x[0]);
            Klasse = x[2];
            Fach = x[3];
            Prüfungsart = x[4];
            PrüfungsartNote = x[5];
            try
            {
                LehrerKürzel = x[7];
            }
            catch (Exception)
            {

                throw;
            }
            
            StudentId = Convert.ToInt32(x[8]);            
        }
                
        private DateTime GetDatum(string datumString)
        {
            return DateTime.ParseExact(datumString, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        }
    }
}