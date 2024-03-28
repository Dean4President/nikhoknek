using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OsztondijtablaGenerator
{
    class NeptunAdatok
    {
        public static List<NeptunAdatok> Adatok = new List<NeptunAdatok>();

        public string NeptunKod { get; private set; }
        public string Kepzes { get; private set; }
        public string HallgatoFelvetel { get; private set; }
        public string Szak { get; private set; }
        public string PenzugyiStatusz { get; private set; }

        public NeptunAdatok(string neptunKod, string kepzes, string hallgatoFelvetel, string szak, string penzugyiStatusz)
        {
            this.NeptunKod = neptunKod.ToUpper();
            this.Kepzes = kepzes;
            this.HallgatoFelvetel = hallgatoFelvetel;
            this.Szak = szak;
            this.PenzugyiStatusz = penzugyiStatusz;
        }

        public static void Beolvas()
        {
            var adatokWorksheet = new XLWorkbook(Config.neptunAdatokFajl).Worksheet(1);

            for (int i = 2; i <= adatokWorksheet.LastRowUsed().RowNumber(); i++)
            {
                var row = adatokWorksheet.Row(i);
                Adatok.Add(new NeptunAdatok(row.Cell(1).GetString().Trim(), row.Cell(2).GetString().Trim(), row.Cell(3).GetString().Trim(), row.Cell(4).GetString().Trim(), row.Cell(5).GetString().Trim()));
            }
        }
    }
}
