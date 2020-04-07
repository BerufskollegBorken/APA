using System.Collections.Generic;

namespace APA
{
    class ZeitUndOrt
    {
        public int Tag;
        public int Stunde;
        public List<string> Raum;

        public ZeitUndOrt(int tag, int stunde, List<string> raum)
        {
            this.Tag = tag;
            this.Stunde = stunde;
            this.Raum = raum;
        }
    }
}