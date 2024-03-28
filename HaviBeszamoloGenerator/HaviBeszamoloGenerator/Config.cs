using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HaviBeszamoloGenerator
{
    public enum Honap
    {
        január = 1, február, március, április, május, június, július, augusztus, szeptember, október, november, december
    }
    public class Config
    {
        public static Honap aktualisHonap = Honap.március;

        public static string tanulmanyiFajl = "adatok/tanulmanyi.xlsx";
        public static string kozeletiFajl = "adatok/kozeleti.xlsx";

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
