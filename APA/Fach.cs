using System.Linq;

namespace APA
{
    public class Fach
    {
        public Fach()
        {
        }

        public Fach(int id, string klasse, string subject, string lehrerkürzel, Noten noten, Sortierung sortierung)
        {
            KürzelUntis = subject;
            Lehrerkürzel = lehrerkürzel;
            Note = (from n in noten
                    where n.Fach == KürzelUntis
                    where n.StudentId == id
                    select n.PrüfungsartNote).FirstOrDefault();

            foreach (var s in sortierung)
            {
                if (s.FachkürzelAtlantis == KürzelUntis)
                {
                    if (klasse.StartsWith(s.Bezeichnung))
                    {
                        Nummer = s.Position1 == 0 ? 99 : s.Position1;
                        break;
                    }
                }
            }            
        }

        public Fach(int nummer, string kürzelUntis, string lernbereich)
        {
            Nummer = nummer;
            KürzelUntis = kürzelUntis;
            Lernbereich = lernbereich;
        }

        public int IdUntis { get; internal set; }
        public string KürzelUntis { get; internal set; }
        public string Lernbereich { get; internal set; }
        public string LangnameUntis { get; internal set; }
        public string BezeichnungImZeugnis { get; internal set; }
        public string Fachklassen { get; internal set; }
        public string Note { get; internal set; }
        public int Nummer { get; private set; }
        public string FachkürzelAtlantis { get; internal set; }
        public string Kurztext { get; internal set; }
        public int Position1 { get; internal set; }
        public int Position2 { get; internal set; }
        public string Bezeichnung { get; internal set; }
        public string Lehrerkürzel { get; private set; }
    }
}