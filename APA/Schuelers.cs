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
            StudentgroupStudents studentgroupStudents = new StudentgroupStudents(exportLessons);
            Noten noten = new Noten();
                        
            Sortierung sortierung = new Sortierung();

            foreach (var schueler in this)
            {                
                // Alle Unterrichte ohne Studentgroup seiner Klasse werden zugeordnet

                foreach (var e in exportLessons)
                {
                    if (e.Klassen.Split('~').Contains(schueler.Klasse.NameUntis)
                                          && e.Teacher != null
                                          && e.Teacher != ""
                                          && e.Subject != null
                                          && e.Subject != ""
                                          && !Global.ZuIgnorierendeFächer.Contains(e.Subject)
                                          && e.StartDate < DateTime.Now
                                          && e.Studentgroup == "")
                    {
                        // Wenn es noch keine Note für das Fach gibt

                        if (!(from n in noten
                              where n.Fach == e.Subject
                              where n.StudentId == schueler.Id
                              select n).Any())
                        {
                            // ... und das Fach mit diesem Lehrer auch noch nicht existiert

                            if (!(from s in schueler.Fächer
                                  where s.Lehrerkürzel == e.Teacher
                                  where s.KürzelUntis == e.Subject
                                  select s).Any())
                            {
                                schueler.Fächer.Add(new Fach(
                                              schueler.Id,
                                              schueler.Klasse.NameUntis,
                                              e.Subject,
                                              e.Teacher,
                                              null,
                                              sortierung
                                          ));
                            }                            
                        }

                            // Wenn es mehr als eine Note für das selbe Fach vom selben Kollegen gibt.
                            foreach (var note in (from n in noten
                                              where n.Fach == e.Subject
                                              where n.StudentId == schueler.Id
                                              select n).ToList())
                        {
                            // ... und es das Fach mit dieser Note noch nicht gibt ...

                            if (!(from s in schueler.Fächer
                                where s.KürzelUntis == e.Subject                                
                                where s.Note == note.PrüfungsartNote
                                select s).Any())
                            {
                                // ... wird es erneut angelegt

                                if (!(from f in schueler.Fächer
                                      where f.KürzelUntis == note.Fach
                                      where f.Lehrerkürzel == note.LehrerKürzel
                                      where f.Note == note.PrüfungsartNote
                                      select f).Any())
                                {
                                    schueler.Fächer.Add(new Fach(
                                              schueler.Id,
                                              schueler.Klasse.NameUntis,
                                              e.Subject,
                                              e.Teacher,
                                              note,
                                              sortierung
                                          ));
                                }                                
                            }                            
                        }
                    }
                }

                // Alle Gruppen werden zu Unterrichten

                foreach (var s in studentgroupStudents)
                {
                    if (s.StudentId == schueler.Id
                                          && s.Subject != null
                                          && s.Subject != ""
                                          && s.StartDate < DateTime.Now)
                    {
                        // Wenn es noch keine Note für das Fach gibt

                        if (!(from n in noten
                              where n.Fach == s.Subject
                              where n.StudentId == schueler.Id
                              select n).Any())
                        {
                            if (!(from f in schueler.Fächer
                                  where f.Lehrerkürzel == (from e in exportLessons
                                                           where e.Studentgroup == s.Studentgroup
                                                           where e.Subject == s.Subject
                                                           where !Global.ZuIgnorierendeFächer.Contains(e.Subject)
                                                           select e.Teacher).FirstOrDefault()
                                  where f.KürzelUntis == s.Subject
                                  select f).Any())
                            {
                                schueler.Fächer.Add(new Fach(
                                              schueler.Id,
                                              schueler.Klasse.NameUntis,
                                              s.Subject,
                                              (from e in exportLessons
                                               where e.Studentgroup == s.Studentgroup
                                               where e.Subject == s.Subject
                                               select e.Teacher).FirstOrDefault(),
                                              null,
                                              sortierung
                                          ));
                            }   
                        }

                        // Wenn es mehr als eine Note für das selbe Fach vom selben Kollegen gibt.
                        foreach (var note in (from n in noten
                                              where n.Fach == s.Subject
                                              where n.StudentId == schueler.Id
                                              select n).ToList())
                        {
                            // ... und es das Fach mit dieser Note noch nicht gibt wird es ...

                            if (!(from f in schueler.Fächer
                                  where f.KürzelUntis == s.Subject
                                  where f.Note == note.PrüfungsartNote
                                  select f).Any())
                            {
                                // ... sofern das Fach mit diesem Kollegen und dieser Note nicht schon existiert ...

                                if (!(from f in schueler.Fächer
                                      where f.KürzelUntis == note.Fach
                                      where f.Lehrerkürzel == note.LehrerKürzel
                                      where f.Note == note.PrüfungsartNote
                                      select f).Any())
                                {
                                    // ... angelegt.

                                    schueler.Fächer.Add(new Fach(
                                                  schueler.Id,
                                                  schueler.Klasse.NameUntis,
                                                  s.Subject,
                                                  note.LehrerKürzel,
                                                  note,
                                                  sortierung
                                              ));
                                }                                
                            }
                        }
                    }
                }

                // Prüfung, ob es noch Noten ohne Unterricht gibt. Das ist der Fall, wenn Fächer nach der Unterstuf nicht fortgeführt wurden.

                foreach (var note in noten)
                {
                    if (note.StudentId == schueler.Id && (note.Fach.StartsWith("BI") || note.Fach.StartsWith("PH")))                                      
                    {
                        if (!(from f in schueler.Fächer where f.KürzelUntis == note.Fach select f).Any())
                        {
                            Console.WriteLine(schueler.Klasse.NameUntis + " " + schueler.Nachname + " hat eine Note in " + note.Fach + " bekommen, aber dieses Fach nicht im Unterricht.");

                            schueler.Fächer.Add(new Fach(
                                                  schueler.Id,
                                                  schueler.Klasse.NameUntis,
                                                  note.Fach + "*",
                                                  note.LehrerKürzel,
                                                  note,
                                                  sortierung
                                              ));
                        }
                    }                    
                }
            }
        }
    }
}   