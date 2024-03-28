using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OsztondijtablaGenerator
{
    class EgyszeriOsztondij : Osztondij
    {
        private const string jogcim = "Egyszeri közéleti ösztöndíj";
        public static string kozteletiTetelNev = "NIK HÖK Közéleti.ö. Egyszeri";
        public static string sportTetelNev = "NIK HÖK Sport.ö. Egyszeri";
        public static string szakmaiTetelNev = "NIK HÖK Szakmai.ö. Egyszeri";
        public static string kulturalisTetelNev = "NIK HÖK Kultúrális.ö. Egyszeri";
        public static string tudomanyosTetelNev = "NIK HÖK Tudományos.ö. Egyszeri";

        public static List<EgyszeriOsztondij> egyszeriOsztondijak = new List<EgyszeriOsztondij>();

        public string Tipus { get; private set; }
        public EgyszeriOsztondij(string neptunKod, string osszeg, string tipus, string indoklas) : base(neptunKod, int.Parse(osszeg), indoklas, jogcim)
        {
            this.Tipus = tipus;

            switch (Tipus)
            {
                case "közéleti":
                    this.TetelNev = $"{kozteletiTetelNev} {Config.aktualisFelev} {Config.aktualisHonap}";
                    break;
                case "sport":
                    this.TetelNev = $"{sportTetelNev} {Config.aktualisFelev} {Config.aktualisHonap}";
                    break;
                case "szakmai":
                    this.TetelNev = $"{szakmaiTetelNev} {Config.aktualisFelev} {Config.aktualisHonap}";
                    break;
                case "kulturalis":
                    this.TetelNev = $"{kulturalisTetelNev} {Config.aktualisFelev} {Config.aktualisHonap}";
                    break;
                case "tudomanyos":
                    this.TetelNev = $"{tudomanyosTetelNev} {Config.aktualisFelev} {Config.aktualisHonap}";
                    break;
                default:
                    break;
            }

        }

        public static void Beolvas()
        {
            EUFETHatarozat();
            Kempelen();
        }

        static void EUFETHatarozat()
        {
            var eufetHatarozatWorksheet = new XLWorkbook(Config.eufetHatarozatFajl).Worksheet(1);

            for (int i = 2; i <= eufetHatarozatWorksheet.LastRowUsed().RowNumber(); i++)
            {
                var row = eufetHatarozatWorksheet.Row(i);

                string neptunKod = row.Cell(2).GetString().Trim();
                string osszeg = row.Cell(5).GetString().Trim();

                if(int.Parse(osszeg) > 0)
                {

                    if (egyszeriOsztondijak.Find(x => neptunKod == x.NeptunKod) is null)
                    {
                        egyszeriOsztondijak.Add(new EgyszeriOsztondij(neptunKod, osszeg, "közéleti", "Érdekvédelmi felelős feladatok ellátása"));
                    }
                    else
                    {
                        egyszeriOsztondijak.Find(x => neptunKod == x.NeptunKod).Osszeg += int.Parse(osszeg);
                    }
                }
            }
        }

        static void Kempelen()
        {
            var egyszeriWorksheet = new XLWorkbook(Config.egyszeriKozeletiFajl).Worksheet(1);

            for (int i = 2; i <= egyszeriWorksheet.LastRowUsed().RowNumber(); i++)
            {
                var row = egyszeriWorksheet.Row(i);

                string neptunKod = row.Cell(1).GetString().Trim();
                string osszeg = row.Cell(5).GetString().Trim();
                string tipus = row.Cell(3).GetString().Trim();
                string indoklas = row.Cell(4).GetString().Trim();

                if(int.Parse(osszeg) > 0)
                {
                    if (egyszeriOsztondijak.Find(x => neptunKod == x.NeptunKod && tipus == x.Tipus) is null)
                    {
                        egyszeriOsztondijak.Add(new EgyszeriOsztondij(neptunKod, osszeg, tipus, indoklas));
                    }
                    else
                    {
                        egyszeriOsztondijak.Find(x => neptunKod == x.NeptunKod && tipus == x.Tipus).Osszeg += int.Parse(osszeg);
                        egyszeriOsztondijak.Find(x => neptunKod == x.NeptunKod && tipus == x.Tipus).Indoklas += $", {indoklas}";
                    }
                }
            }
        }
    }
}
