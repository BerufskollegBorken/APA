using System.Collections.Generic;

namespace APA
{
    public class Stundentafel
    {
        public Stundentafel()
        {
            Fachs = new List<Fach>();
        }

        public int IdUntis { get; internal set; }
        public string Name { get; internal set; }
        public string Langname { get; internal set; }
        public List<Fach> Fachs { get; internal set; }
    }
}