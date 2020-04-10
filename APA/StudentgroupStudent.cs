using Microsoft.Office.Interop.Excel;
using System;
using System.Globalization;

namespace APA
{
    public class StudentgroupStudent
    {
        public int StudentId { get; private set; }
        public string Forename { get; private set; }
        public string Name { get; private set; }
        public string Studentgroup { get; private set; }
        public string Subject { get; private set; }
        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }

        public StudentgroupStudent(string line)
        {
            var x = line.Split('\t');
            StudentId = Convert.ToInt32(x[0]);
            Name = x[1];
            Forename = x[2];
            Studentgroup = x[3];
            Subject = x[4];
            StartDate = GetDatum(x[5]);
            EndDate = GetDatum(x[6]);
        }
                
        private DateTime GetDatum(string datumString)
        {
            if (datumString == "")
            {
                return new DateTime();
            }
            return DateTime.ParseExact(datumString, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        }
    }
}