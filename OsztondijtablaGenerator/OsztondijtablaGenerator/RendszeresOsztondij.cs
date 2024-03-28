using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OsztondijtablaGenerator
{
    class RendszeresOsztondij : Osztondij
    {
        private const string jogcim = "Rendszeres közéleti ösztöndíj";
        public static string tetelNev = "NIK HÖK Közéleti.ö. Rendszeres";

        public static List<RendszeresOsztondij> RendszeresOsztondijak = new List<RendszeresOsztondij>();

        public string TetelNev { get; private set; }

        public RendszeresOsztondij(string neptunKod, string osszeg, string indoklas) : base(neptunKod, int.Parse(osszeg), indoklas, jogcim)
        {
            this.TetelNev = $"{tetelNev} {Config.aktualisFelev} {Config.aktualisHonap}";
        }

        public static void Beolvas()
        {
            var elnokiHatarozatWorksheet = new XLWorkbook(Config.elnokiHatarozatFajl).Worksheet(1);

            for (int i = 2; i <= elnokiHatarozatWorksheet.LastRowUsed().RowNumber(); i++)
            {
                var row = elnokiHatarozatWorksheet.Row(i);
                RendszeresOsztondijak.Add(new RendszeresOsztondij(row.Cell(1).GetString().Trim(), row.Cell(2).GetString().Trim(), row.Cell(3).GetString().Trim()));
            }
        }
    }
}
