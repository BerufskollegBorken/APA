using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace APA
{
    public class Schueler
    {
        /// <summary>
        /// Atlantis-ID
        /// </summary>
        public int Id { get; set; }
        public string Nachname { get; private set; }
        public string Vorname { get; private set; }
        public Klasse Klasse { get; private set; }
        public DateTime Gebdat { get; set; }
        public bool IstVolljährig
        {
            get
            {
                if (DateTime.Now >= Gebdat.AddYears(18))
                {
                    return true;
                }
                return false;
            }
        }
        public bool IstSchulpflichtig
        {
            get
            {
                try
                {
                    // Minderjährige sind schulpflichtig

                    if (DateTime.Now < Gebdat.AddYears(18))
                    {
                        return true;
                    }

                    // Wenn ein Vollzeitschüler ..

                    if (!Klasse.Jahrgang.StartsWith("BS"))
                    {
                        // ... 18 ist ...

                        if (DateTime.Now >= Gebdat.AddYears(18))
                        {
                            // ...  aber erst nach SJ-Beginn 18 geworden ist, ...

                            if (Gebdat.AddYears(18) >= (new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1), 8, 1)))
                            {
                                // ... dann ist er bis zum Ende des SJ schulpflichtig.

                                return true;
                            }
                        }
                    }

                    // Wenn ein Berufsschüler ...

                    if (Klasse.Jahrgang.StartsWith("BS"))
                    {
                        // ... vor der Vollendung seines 21. Lebensjahrs die Berufsausbildung beginnt, ...

                        if (Bildungsgangeintrittsdatum < Gebdat.AddYears(21))
                        {
                            // ... ist er bis zum Ende berufsschulpflichtig.

                            return true;
                        }
                    }
                }
                catch (Exception)
                {
                    return false;
                }
                return false;
            }
        }

        /// <summary>
        /// Jede Abwesenheit steht für das Fehlen eines Schülers an einem Schultag
        /// </summary>
        public List<Note> Noten { get; private set; }
        public int FehltUnunterbrochenUnentschuldigtSeitTagen { get; set; }

        public List<Maßnahme> Maßnahmen { get; set; }

        public Adresse Adresse { get; private set; }
        
        public DateTime Bildungsgangeintrittsdatum { get; private set; }
        public Feriens Feriens { get; private set; }
        public int AktSj { get; private set; }
        public Unterrichte Unterrichte { get; internal set; }
        public List<Fach> Fächer { get; internal set; }

        public Schueler(int id, string nachname, string vorname, DateTime gebdat, Klasse klasse, DateTime bildungsgangeintrittsdatum)
        {
            Id = id;
            Nachname = nachname;
            Vorname = vorname;
            Klasse = klasse;
            Gebdat = gebdat;
            Bildungsgangeintrittsdatum = bildungsgangeintrittsdatum;
            Noten = new List<Note>();
            Maßnahmen = new List<Maßnahme>();
            Fächer = new List<Fach>();
        }

        internal string Render(string m)
        {
            /* var x = (from o in Maßnahmen where o.Kürzel == m select o).FirstOrDefault();

             if (x != null)
             {
                 var z = (from aaa in x.AngemahnteAbwesenheitenDieserMaßnahme select aaa.Fehlstunden).Sum();

                 return x.Datum.ToShortDateString() + "(" + z + ")";
             }*/
            return "";
        }

        internal void RenderMaßnahmen()
        {
            foreach (var om in this.Maßnahmen)
            {
                Console.WriteLine("      " + om.Beschreibung + " (" + om.Datum.ToShortDateString() + ")");
            }
        }

        internal string GetE1Datum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "E1" select o).Any())
            {
                return (from o in this.Maßnahmen where o.Kürzel == "E1" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        internal void GetAdresse(string aktSjAtlantis, string connectionStringAtlantis)
        {
            using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
            {
                DataSet dataSet = new DataSet();
                OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT DBA.adresse.pu_id AS ID,
DBA.adresse.plz AS PLZ,
DBA.adresse.ort AS Ort,
DBA.adresse.strasse AS Strasse
FROM DBA.adresse
WHERE ID = " + Id + " AND hauptadresse_jn = 'j'", connection);

                connection.Open();
                schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                {
                    int id = Convert.ToInt32(theRow["ID"]);
                    string plz = theRow["PLZ"] == null ? "" : theRow["PLZ"].ToString();
                    string ort = theRow["Ort"] == null ? "" : theRow["Ort"].ToString();
                    string strasse = theRow["Strasse"] == null ? "" : theRow["Strasse"].ToString();

                    Adresse adresse = new Adresse(
                        id,
                        plz,
                        ort,
                        strasse)
                        ;

                    this.Adresse = adresse;
                }

                connection.Close();
            }
        }

        internal string GetADatum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "A" select o).Any())
            {
                return (from o in this.Maßnahmen where o.Kürzel == "A" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }
    }
}