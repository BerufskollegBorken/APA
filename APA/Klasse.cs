using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
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
            var ms = new MemoryStream();
            TextWriter tw = new StreamWriter(ms);

            Worksheet deckblatt = workbook.Worksheets.get_Item(1);
            deckblatt.Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            workbook.Sheets[workbook.Sheets.Count].Name = NameUntis + "-D";
            var worksheet = workbook.Sheets[NameUntis + "-D"];
            worksheet.Activate();

            worksheet.Cells[7, 2] = "Prüfung: Sommer " + DateTime.Now.Year;
            worksheet.Cells[10, 3] = NameUntis;
            worksheet.Cells[10, 6] = Klassenleitungen[0].Vorname + " " + Klassenleitungen[0].Nachname;
            tw.WriteLine("Prüfung: Sommer " + DateTime.Now.Year);
            tw.WriteLine(NameUntis);

            int z = 16;

            // Lehrer auf dem Deckblatt auflisten

            foreach (var lehrerkürzel in (from s in schuelers from f in s.Fächer orderby f.Lehrerkürzel select f.Lehrerkürzel).Distinct())
            {
                worksheet.Cells[z, 2] = (from l in lehrers where l.Kürzel == lehrerkürzel select l.Nachname + ", " + l.Vorname).FirstOrDefault();

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
            //worksheet.PageSetup.CenterHeader = "Abschlusskonferenz";
            //worksheet.PageSetup.RightHeader = DateTime.Now.ToLocalTime();

            worksheet.Cells[1, 1] = "Klasse: " + this.NameUntis;
            worksheet.Cells[1, 4] = "Klassenleitung: " + this.Klassenleitungen[0].Vorname + " " + this.Klassenleitungen[0].Nachname + " " + this.Klassenleitungen[0].Mail;
            worksheet.Cells[1, 19] = "Schuljahr: " + Global.AktSjAtl;
            worksheet.Cells.Font.Size = 12;

            int zeileObenLinks = 3;
            int spalteObenlinks = 1;

            foreach (var schueler in schuelers.OrderBy(x => x.Nachname).ThenBy(y => y.Vorname).ToList())
            {
                worksheet.Cells[zeileObenLinks + 2, spalteObenlinks] = schueler.Nachname + ", " + schueler.Vorname;
                worksheet.Cells[zeileObenLinks + 3, spalteObenlinks] = "*" + schueler.Gebdat.ToShortDateString();

                if (NameUntis == "BSO")
                {
                    worksheet.Cells[zeileObenLinks + 6, spalteObenlinks + 1] = "";
                    worksheet.Cells[zeileObenLinks + 7, spalteObenlinks + 1] = "";
                }

                int x = 0;

                foreach (var fach in (from f in schueler.Fächer where !f.KürzelUntis.EndsWith("FU") select f).OrderBy(y => y.Nummer).ToList())
                {
                    worksheet.Cells[zeileObenLinks + 2, spalteObenlinks + 2 + x] = fach.KürzelUntis;
                    worksheet.Cells[zeileObenLinks + 3, spalteObenlinks + 2 + x] = fach.Note;
                    x++;
                }
                zeileObenLinks = zeileObenLinks + 12;
            }

            Console.Write("Excel-Worksheet " + 12M + " nach PDF umwandeln");
            worksheet.ExportAsFixedFormat(
                XlFixedFormatType.xlTypePDF,
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + NameUntis,
                XlFixedFormatQuality.xlQualityStandard,
                true,
                true,
                1,
                10,
                false);

            Console.WriteLine(" ... ok");
            Console.Write(NameUntis + ".pdf kompremieren");
            CompressDirectory(
                new List<string>() { Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + NameUntis + ".pdf" },
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\");
            Console.WriteLine(" ... ok");
            Console.Write(" Mail senden " + NameUntis);

            string kla = "";

            foreach (var item in Klassenleitungen)
            {
                kla += item.Vorname + " " + item.Nachname + ",";
            }

            Global.MailSenden(
                this, 
                "Notenliste " + NameUntis + " für " + kla ,
                @"Guten Morgen " + kla + @"<br><br>
zur Vorbereitung auf die Zulasungskonferenz der Klasse " + NameUntis + @" am 21.4.20 erhalten Sie die Liste der Noten Ihrer Klasse.
<br>
<br>
Ihre Aufgabe ist es, fehlende Noten bei den Fachkolleginnen und Fachkollegen anzufordern. Die Noten müssen dann von der jeweiligen Lehrkraft bis spätestens 13.4.20 im DigiKlas eingetragen werden. Parallel zu Ihren Bemühungen werden automatisch Mails mit der Aufforderung zur Eintragung  verschickt. Am 13.4.20 um 24 Uhr wird die Änderung- / Neueingabemöglichkeit gesperrt.
<br>
Es werden Ihnen in der Liste alle Fächer angezeigt, die seit dem Schuljahresbeginn unterricht wurden. Das schließt auch diejenigen Fächer ein, die z.B. in der zweiten Woche nach den Ferien ersatzlos gestrichen wurden. Als Klassenleitung wissen Sie, wo entsprechend keine Noten erforderlich sind und wo noch Noten fehlen.

Aus Datenschutzgründen kann die Liste natürlich nicht unverschlüsselt gesendet werden. Das Kennwort ist unsere leicht abgewandelte Schulnummer. Sie finden das Kennwort <a href='https://bk-borken.lms.schulon.org/course/view.php?id=415'>hier</a>. <br><br>Frohe Ostern<br><br>Stefan Bäumer<br><br>PS: Weil diese Mail samt Inhalt automatisch erstellt und versandt wurde, ist der (angekündigte) Versand der Liste über den Messenger so nicht möglich.", new List<string>() { Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + NameUntis + ".zip" });

            Console.WriteLine(" ... ok");
        }

        public void CompressDirectory(List<string> filenames, string OutputFilePath, int CompressionLevel = 9)
        {
            try
            {
                using (ZipOutputStream OutputStream = new ZipOutputStream(File.Create(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + NameUntis + ".zip")))
                {
                    OutputStream.SetLevel(CompressionLevel);
                    OutputStream.Password = "!7765Neun";

                    byte[] buffer = new byte[4096];

                    foreach (string file in filenames)
                    {
                        ZipEntry entry = new ZipEntry(Path.GetFileName(file))
                        {
                            DateTime = DateTime.Now
                        };
                        OutputStream.PutNextEntry(entry);

                        using (FileStream fs = File.OpenRead(file))
                        {
                            int sourceBytes;

                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                OutputStream.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }
                    OutputStream.Finish();
                    OutputStream.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception during processing {0}", ex);
            }
        }
    }
}