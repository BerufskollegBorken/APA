namespace APA
{
    public class Raum
    {
        public string RaumApa { get; set; }

        public Raum(string raumApa)
        {
            RaumApa = raumApa;
        }

        public string Raumname { get; internal set; }
    }
}