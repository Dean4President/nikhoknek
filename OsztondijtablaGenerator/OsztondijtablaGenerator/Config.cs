using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OsztondijtablaGenerator
{
    public enum Honap
    {
        január = 1, február, március, április, május, június, július, augusztus, szeptember, október, november, december
    }
    public class Config
    {
        public static Honap aktualisHonap = Honap.március;
        public static string aktualisFelev = "2023/24/2";

        public static string neptunAdatokFajl = "adatok/neptun_adatok.xlsx";
        public static string elnokiHatarozatFajl = "adatok/elnoki_hatarozat.xlsx";
        public static string eufetHatarozatFajl = "adatok/eufet_hatarozat.xlsx";
        public static string egyszeriKozeletiFajl = "adatok/leadottak.xlsx";

        public static string AktualisHonap
        {
            get
            {
                int honap = (int)aktualisHonap;
                return honap < 10 ? $"0{honap}" : $"{honap}";
            }
        }
    }
}
