using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OsztondijtablaGenerator
{
    abstract class Osztondij
    {
        public string NeptunKod { get; private set; }
        public int Osszeg { get; set; }
        public string TetelNev { get; set; }
        public string Indoklas { get; set; }
        public string Jogcim { get; set; }

        protected Osztondij(string neptunKod, int osszeg, string indoklas, string jogcim)
        {
            this.NeptunKod = neptunKod.ToUpper();
            this.Osszeg = osszeg;
            this.Indoklas = indoklas;
            this.Jogcim = jogcim;
        }

        public static void Generalas()
        {
            NeptunAdatok.Beolvas();
        }
    }
}
