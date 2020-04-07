using System;
using System.Globalization;

namespace APA
{
    public class ExportLesson
    {
        private string line;

        public ExportLesson(string line)
        {
            var x = line.Split('\t');
            LessonId = Convert.ToInt32(x[0]);
            LessonNumber = Convert.ToInt32(x[1]);
            Subject = x[2];
            Teacher = x[3];
            Klassen = x[4];
            Studentgroup = x[5];
            Periods = x[6];
            StartDate = GetDatum(x[7]);
            EndDate = GetDatum(x[8]);
            Room = x[9];
        }

        public int LessonId { get; private set; }
        public int LessonNumber { get; private set; }
        public string Subject { get; private set; }
        public string Teacher { get; private set; }
        public string Klassen { get; private set; }
        public string Studentgroup { get; private set; }
        public string Periods { get; private set; }
        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }
        public string Room { get; private set; }
        public int ForeignKey { get; private set; }

        private DateTime GetDatum(string datumString)
        {
            return DateTime.ParseExact(datumString, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        }
    }
}