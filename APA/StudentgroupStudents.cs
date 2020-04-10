using System;
using System.Collections.Generic;
using System.IO;

namespace APA
{
    public class StudentgroupStudents : List<StudentgroupStudent>
    {
        public StudentgroupStudents()
        {
            using (StreamReader reader = new StreamReader(Global.InputStudentgroupStudentsCsv))
            {
                string überschrift = reader.ReadLine();

                while (true)
                {
                    string line = reader.ReadLine();

                    if (line != null)
                    {
                        StudentgroupStudent studentgroupStudent = new StudentgroupStudent(line);
                        this.Add(studentgroupStudent);
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine(("StudentgroupStudents " + ".".PadRight(this.Count / 500, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');
                
            }
        }
    }
}