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
        public Raum Raum { get; internal set; }
        public string Url { get; internal set; }
        public string Jahrgang { get; internal set; }
        public DateTime ErsterSchultag { get; internal set; }
        
        internal void Notenliste(Application application, Workbook workbook, List<Schueler> schuelers)
        {
            Worksheet vorlage = workbook.Worksheets.get_Item(1);

            Range kopfzeile = vorlage.get_Range("A1:V2");
            Range rangeSchueler = vorlage.get_Range("A2:V14");

            Worksheet worksheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            //Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = this.NameUntis;            
            worksheet.Cells[1, 1] = "Klasse: " + this.NameUntis;
            worksheet.Cells[1, 4] = "Klassenleitung: " + this.Klassenleitungen[0].Vorname + " " + this.Klassenleitungen[0].Nachname;
            worksheet.Cells[1, 7] = "Schuljahr: " + Global.AktSjAtl;
            worksheet.Cells.Font.Size = 12;
                        
            Range to = worksheet.get_Range("A1:V2");
            kopfzeile.Copy(to);
                        
            int zeileObenLinks = 2;
            int spalteObenlinks = 1;

            foreach (var schueler in schuelers.OrderBy(x => x.Nachname).ThenBy(y => y.Vorname).ToList())
            {
                to = worksheet.get_Range("A" + zeileObenLinks + ":V" + (zeileObenLinks + 12));

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
            
            Console.WriteLine(NameUntis + " ... ok");
        }
    }
}