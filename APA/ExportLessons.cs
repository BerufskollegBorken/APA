using System;
using System.Collections.Generic;
using System.IO;

namespace APA
{
    public class ExportLessons : List<ExportLesson>
    {
        public ExportLessons()
        {
            using (StreamReader reader = new StreamReader(Global.InputExportLessons))
            {
                string überschrift = reader.ReadLine();

                while (true)
                {
                    string line = reader.ReadLine();

                    if (line != null)
                    {
                        ExportLesson exportLesson = new ExportLesson(line);
                        
                        this.Add(exportLesson);
                                                
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine(("ExportLessons " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');
            }
        }
    }
}