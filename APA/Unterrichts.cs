using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace APA
{
    public class Unterrichts : List<Unterricht>
    {
        public Lehrers Lehrers { get; set; }
        public DateTime Schuljahresbeginn { get; private set; }

        public Unterrichts()
        {
        }

        public Unterrichts(int periode, Klasses klasses, Lehrers lehrers, Fachs fachs, Raums raums)
        {
            Lehrers = lehrers;

            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConU))
            {
                int id = 0;

                try
                {
                    string queryString = @"SELECT DISTINCT 
Lesson_ID,
LessonElement1,
Periods,
Lesson.LESSON_GROUP_ID,
Lesson_TT,
Flags,
DateFrom,
DateTo
FROM LESSON
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSjUnt + ") AND ((TERM_ID)=" + periode + ") AND ((Lesson.SCHOOL_ID)=177659) AND (((Lesson.Deleted)=No))) ORDER BY LESSON_ID;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    Console.WriteLine("Unterrichte");
                    Console.WriteLine("-----------");

                    while (oleDbDataReader.Read())
                    {
                        id = oleDbDataReader.GetInt32(0);

                        string wannUndWo = Global.SafeGetString(oleDbDataReader, 4);

                        var zur = wannUndWo.Replace("~~", "|").Split('|');

                        ZeitUndOrts zeitUndOrts = new ZeitUndOrts();

                        for (int i = 0; i < zur.Length; i++)
                        {
                            if (zur[i] != "")
                            {
                                var zurr = zur[i].Split('~');

                                int tag = 0;
                                int stunde = 0;
                                List<string> raum = new List<string>();

                                try
                                {
                                    tag = Convert.ToInt32(zurr[1]);
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Der Unterricht " + id + " hat keinen Tag.");
                                }

                                try
                                {
                                    stunde = Convert.ToInt32(zurr[2]);
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Der Unterricht " + id + " hat keine Stunde.");
                                }

                                try
                                {
                                    var ra = zurr[3].Split(';');

                                    foreach (var item in ra)
                                    {
                                        if (item != "")
                                        {
                                            raum.AddRange((from r in raums
                                                           where item.Replace(";", "") == r.IdUntis.ToString()
                                                           select r.Raumnummer));
                                        }
                                    }

                                    if (raum.Count == 0)
                                    {
                                        raum.Add("");
                                    }
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Der Unterricht " + id + " hat keinen Raum.");
                                }


                                ZeitUndOrt zeitUndOrt = new ZeitUndOrt(tag, stunde, raum);
                                zeitUndOrts.Add(zeitUndOrt);
                            }
                        }

                        string lessonElement = Global.SafeGetString(oleDbDataReader, 1);

                        int anzahlGekoppelterLehrer = lessonElement.Count(x => x == '~') / 21;

                        List<string> klassenKürzel = new List<string>();

                        for (int i = 0; i < anzahlGekoppelterLehrer; i++)
                        {
                            var lesson = lessonElement.Split(',');

                            var les = lesson[i].Split('~');
                            string lehrer = les[0] == "" ? null : (from l in lehrers where l.IdUntis.ToString() == les[0] select l.Kürzel).FirstOrDefault();

                            string fach = les[2] == "0" ? "" : (from f in fachs where f.IdUntis.ToString() == les[2] select f.KürzelUntis).FirstOrDefault();

                            string raumDiesesUnterrichts = "";
                            if (les[3] != "")
                            {
                                raumDiesesUnterrichts = (from r in raums where (les[3].Split(';')).Contains(r.IdUntis.ToString()) select r.Raumnummer).FirstOrDefault();
                            }

                            int anzahlStunden = oleDbDataReader.GetInt32(2);

                            if (les.Count() >= 17)
                            {
                                foreach (var kla in les[17].Split(';'))
                                {
                                    Klasse klasse = new Klasse();

                                    if (kla != "")
                                    {
                                        if (!(from kl in klassenKürzel
                                              where kl == (from k in klasses
                                                           where k.IdUntis == Convert.ToInt32(kla)
                                                           select k.NameUntis).FirstOrDefault()
                                              select kl).Any())
                                        {
                                            klassenKürzel.Add((from k in klasses
                                                               where k.IdUntis == Convert.ToInt32(kla)
                                                               select k.NameUntis).FirstOrDefault());
                                        }
                                    }
                                }
                            }
                            else
                            {
                            }

                            if (lehrer != null)
                            {
                                for (int z = 0; z < zeitUndOrts.Count; z++)
                                {
                                    // Wenn zwei Lehrer gekoppelt sind und zwei Räume zu dieser Stunde gehören, dann werden die Räume entsprechend verteilt.

                                    string r = zeitUndOrts[z].Raum[0];
                                    try
                                    {
                                        if (anzahlGekoppelterLehrer > 1 && zeitUndOrts[z].Raum.Count > 1)
                                        {
                                            r = zeitUndOrts[z].Raum[i];
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        if (anzahlGekoppelterLehrer > 1 && zeitUndOrts[z].Raum.Count > 1)
                                        {
                                            r = zeitUndOrts[z].Raum[0];
                                        }
                                    }

                                    string k = "";

                                    foreach (var item in klassenKürzel)
                                    {
                                        k += item + ",";
                                    }

                                    // Nur wenn der tagDesUnterrichts innerhalb der Befristung stattfindet, wird er angelegt

                                    DateTime von = DateTime.ParseExact((oleDbDataReader.GetInt32(6)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                    DateTime bis = DateTime.ParseExact((oleDbDataReader.GetInt32(7)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);


                                    Unterricht unterricht = new Unterricht(
                                        id,
                                        lehrer,
                                        fach,
                                        k.TrimEnd(','),
                                        r,
                                        "",
                                        zeitUndOrts[z].Tag,
                                        zeitUndOrts[z].Stunde);

                                    this.Add(unterricht);
                                }
                            }
                        }
                    }
                    Console.WriteLine(("Unterrichte " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

                    oleDbDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Fehler beim Unterricht mit der ID " + id + "\n" + ex.ToString());
                    throw new Exception("Fehler beim Unterricht mit der ID " + id + "\n" + ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

    }
}