using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace APA
{
    public class Klasse
    {
        public int IdUntis { get; internal set; }
        public string NameUntis { get; internal set; }
        public List<Lehrer> Klassenleitungen { get; internal set; }
        public string Bereichsleitung { get; internal set; }
        public string Beschreibung { get; internal set; }
        public string Url { get; internal set; }
        public string Jahrgang { get; internal set; }
        public DateTime ErsterSchultag { get; internal set; }
        
        internal void Notenliste(Application application, Workbook workbook, List<Schueler> schuelers, Lehrers lehrers)
        {
            Worksheet deckblatt = workbook.Worksheets.get_Item(1);
            deckblatt.Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            workbook.Sheets[workbook.Sheets.Count].Name = NameUntis + "-D";
            var worksheet = workbook.Sheets[NameUntis + "-D"];
            worksheet.Activate();

            worksheet.Cells[7, 2] = "Prüfung: Sommer " + DateTime.Now.Year;
            worksheet.Cells[10, 3] = NameUntis;
            worksheet.Cells[10, 6] = Klassenleitungen[0].Vorname + " " +Klassenleitungen[0].Nachname;
            int z = 16;
            foreach (var lehrerkürzel in (from s in schuelers from f in s.Fächer orderby f.Lehrerkürzel select f.Lehrerkürzel).Distinct())
            {
                worksheet.Cells[z, 2] = (from l in lehrers where l.Kürzel == lehrerkürzel select l.Nachname + ", " + l.Vorname ).FirstOrDefault();

                var fächer = (from s in schuelers
                              from f in s.Fächer
                              where f.Lehrerkürzel == lehrerkürzel
                              where !f.KürzelUntis.EndsWith(" FU")
                              select f.KürzelUntis).Distinct().ToList();
                var ff = "";
                foreach (var fach in fächer)
                {
                    ff += fach + ",";
                }
                worksheet.Cells[z, 6] = ff.TrimEnd(',');
                z++;
            }
            
            Worksheet vorlage = workbook.Sheets["Liste"];
            
            vorlage.Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            workbook.Sheets[workbook.Sheets.Count].Name = NameUntis + "-L";
            worksheet = workbook.Sheets[NameUntis + "-L"];
            worksheet.Activate();

            worksheet.PageSetup.LeftHeader = "Prüfungsliste";
            worksheet.PageSetup.CenterHeader = "Abschlusskonferenz";
            worksheet.PageSetup.RightHeader = DateTime.Now.ToLocalTime();

            worksheet.Cells[1, 1] = "Klasse: " + this.NameUntis;
            worksheet.Cells[1, 4] = "Klassenleitung: " + this.Klassenleitungen[0].Vorname + " " + this.Klassenleitungen[0].Nachname;            
            worksheet.Cells[1, 12] = "Schuljahr: " + Global.AktSjAtl;
            worksheet.Cells.Font.Size = 12;
                        
            int zeileObenLinks = 3;
            int spalteObenlinks = 1;

            foreach (var schueler in schuelers.OrderBy(x => x.Nachname).ThenBy(y => y.Vorname).ToList())
            {
                worksheet.Cells[zeileObenLinks + 2, spalteObenlinks] = schueler.Nachname + ", " + schueler.Vorname;
                worksheet.Cells[zeileObenLinks + 3 , spalteObenlinks] = "*" + schueler.Gebdat.ToShortDateString();

                int x = 0;

                foreach (var fach in (from f in schueler.Fächer where !f.KürzelUntis.EndsWith("FU") select f).OrderBy(y=>y.Nummer).ToList())
                {
                    worksheet.Cells[zeileObenLinks + 2, spalteObenlinks + 2 + x] = fach.KürzelUntis;
                    worksheet.Cells[zeileObenLinks + 3, spalteObenlinks + 2 + x] = fach.Note;
                    x++;
                }
                zeileObenLinks = zeileObenLinks + 12;
            }

            // Verbleibende Zellen Löschen

            //Range TempRange = worksheet.get_Range("A" + zeileObenLinks, "V450");

            // 1. To Delete Entire Row - below rows will shift up
            //TempRange.EntireRow.Delete(Type.Missing);

            // 2. To Delete Cells - Below cells will shift up
            //TempRange.Cells.Delete(Type.Missing);

            Console.WriteLine(NameUntis + " ... ok");
        }
    }
}