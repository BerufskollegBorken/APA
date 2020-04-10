using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace APA
{
    public class Schuelers : List<Schueler>
    {
        public Schuelers(Klasses klss, Lehrers lehs)
        {
            using (OdbcConnection connection = new OdbcConnection(Global.ConAtl))
            {
                DataSet dataSet = new DataSet();
                OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.schue_sj.pu_id AS ID,
DBA.schue_sj.dat_eintritt AS bildungsgangEintrittDatum,
DBA.schue_sj.dat_austritt AS Austrittsdatum,
DBA.schue_sj.s_klassenziel_erreicht,
DBA.schue_sj.dat_klassenziel_erreicht,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt AS GebDat,
DBA.klasse.klasse AS Klasse
FROM ( DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id ) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id 
WHERE vorgang_schuljahr = '" + Global.AktSjAtl + "'", connection);

                connection.Open();
                schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                {
                    int id = Convert.ToInt32(theRow["ID"]);

                    DateTime austrittsdatum = theRow["Austrittsdatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Austrittsdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                    DateTime bildungsgangEintrittDatum = theRow["bildungsgangEintrittDatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["bildungsgangEintrittDatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                    if (austrittsdatum.Year == 1)
                    {
                        DateTime gebdat = theRow["Gebdat"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Gebdat"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                        Klasse klasse = theRow["Klasse"] == null ? null : (from k in klss where k.NameUntis == theRow["Klasse"].ToString() select k).FirstOrDefault();

                        string nachname = theRow["Nachname"] == null ? "" : theRow["Nachname"].ToString();
                        string vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();

                        Schueler schueler = new Schueler(
                            id,
                            nachname,
                            vorname,
                            gebdat,
                            klasse,
                            bildungsgangEintrittDatum
                            );

                        if ((from a in Global.AbschlussKlassen
                             where klasse != null
                             where klasse.NameUntis != null
                             where klasse.NameUntis.StartsWith(a)
                             select a).Any())
                        {
                            this.Add(schueler);
                        }
                    }
                }

                connection.Close();
                Console.WriteLine(("Schüler " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');
            }
        }

        internal void Unterrichte()
        {
            ExportLessons exportLessons = new ExportLessons();
            StudentgroupStudents studentgroupStudents = new StudentgroupStudents();
            Noten noten = new Noten();
            Sortierung sortierung = new Sortierung();

            foreach (var schueler in this)
            {
                // Alle Unterrichte ohne Studentgroup seiner Klasse werden zugeordnet
                schueler.Fächer.AddRange((from e in exportLessons
                                          where (e.Klassen.Split('~')).Contains(schueler.Klasse.NameUntis)
                                          where e.Studentgroup == ""
                                          select new Fach
                                          (
                                              schueler.Id,
                                              schueler.Klasse.NameUntis,
                                              e.Subject,
                                              e.Teacher,
                                              noten, 
                                              sortierung
                                          )
                                              
                                              ).ToList());

                // Alle Gruppen werden zu Unterrichten
                schueler.Fächer.AddRange((from s in studentgroupStudents
                                          where s.StudentId == schueler.Id                                          
                                          select new Fach
                                          (
                                              schueler.Id,
                                              schueler.Klasse.NameUntis,
                                              s.Subject,
                                              (from e in exportLessons
                                               where e.Studentgroup == s.Studentgroup
                                               select e.Teacher).FirstOrDefault(),
                                              noten,
                                              sortierung
                                          )).ToList());



            }
        }
    }
}   