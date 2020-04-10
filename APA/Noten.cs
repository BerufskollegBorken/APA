using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace APA
{
    public class Noten : List<Note>
    {
        public Noten()
        {
            using (StreamReader reader = new StreamReader(Global.InputNotenCsv))
            {
                string überschrift = reader.ReadLine();
                int zeile = 1;

                while (true)
                {
                    string line = reader.ReadLine();
                    
                    if (line != null)
                    {
                        Note note = new Note(line, zeile);
                        
                        if (note.Prüfungsart == "Jahreszeugnis")
                        {
                            if ((from a in Global.AbschlussKlassen
                                 where note.Klasse != null
                                 where note.Klasse != null
                                 where note.Klasse.StartsWith(a)
                                 select a).Any())
                            {
                                this.Add(note);
                            }
                        }
                        zeile++;
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine(("Noten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');
            }
        }
    }
}