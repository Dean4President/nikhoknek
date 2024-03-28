using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OsztondijtablaGenerator
{
    class Tetel
    {
        public static List<Tetel> tetelek = new List<Tetel>();

        public string Owner { get; private set; }
        public string Kepzes { get; private set; }
        public string Felev { get; private set; }
        public int Value { get; private set; }
        public string TetelNev { get; private set; }
        public string HallgatoFelvetele { get; private set; }
        public string Szak { get; private set; }
        public string PenzugyiStatusz { get; private set; }

        public string Indoklas { get; private set; }
        public string Jogcim { get; private set; }

        public Tetel(string owner, string kepzes, string felev, int value, string tetelNev, string hallgatoFelvetele, string szak, string penzugyiStatusz, string indoklas, string jogcim)
        {
            this.Owner = owner;
            this.Kepzes = kepzes;
            this.Felev = felev;
            this.Value = value;
            this.TetelNev = tetelNev;
            this.HallgatoFelvetele = hallgatoFelvetele;
            this.Szak = szak;
            this.PenzugyiStatusz = penzugyiStatusz;
            this.Indoklas = indoklas;
            this.Jogcim = jogcim;
        }

        public static void TetelGeneralas()
        {
            foreach (RendszeresOsztondij item in RendszeresOsztondij.RendszeresOsztondijak)
            {
                NeptunAdatok aktualis = NeptunAdatok.Adatok.Find(a => item.NeptunKod == a.NeptunKod);
                
                tetelek.Add(new Tetel(item.NeptunKod, aktualis.Kepzes, Config.aktualisFelev, item.Osszeg, item.TetelNev, aktualis.HallgatoFelvetel, aktualis.Szak, aktualis.PenzugyiStatusz, item.Indoklas, item.Jogcim));
            }

            foreach (EgyszeriOsztondij item in EgyszeriOsztondij.egyszeriOsztondijak)
            {
                NeptunAdatok aktualis = NeptunAdatok.Adatok.Find(a => item.NeptunKod == a.NeptunKod);

                tetelek.Add(new Tetel(item.NeptunKod, aktualis.Kepzes, Config.aktualisFelev, item.Osszeg, item.TetelNev, aktualis.HallgatoFelvetel, aktualis.Szak, aktualis.PenzugyiStatusz, item.Indoklas, item.Jogcim));
            }

            TablaGeneralas();
        }

        static void TablaGeneralas()
        {
            KozeletiNyers();
            Osszesito();
        }

        static void KozeletiNyers()
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Munka1");
            worksheet.Cell(1, 1).Value = "OWNER";
            worksheet.Cell(1, 2).Value = "KEPZES";
            worksheet.Cell(1, 3).Value = "FELEV";
            worksheet.Cell(1, 4).Value = "VALUE";
            worksheet.Cell(1, 5).Value = "PENZUGYIKOD";
            worksheet.Cell(1, 6).Value = "PENZUGYIAZONOSITO";
            worksheet.Cell(1, 7).Value = "KIFIZETESITETELNEVE";
            worksheet.Cell(1, 8).Value = "SZAK";
            worksheet.Cell(1, 9).Value = "HALLGATOFELVETEL";
            worksheet.Cell(1, 10).Value = "SZAK";
            worksheet.Cell(1, 11).Value = "Pénzügyi státusz";

            int row = 2;
            foreach (Tetel tetel in tetelek.OrderBy(x => x.Owner))
            {
                worksheet.Cell(row, 1).Value = tetel.Owner;
                worksheet.Cell(row, 2).Value = tetel.Kepzes;
                worksheet.Cell(row, 3).Value = Config.aktualisFelev;
                worksheet.Cell(row, 4).Value = tetel.Value;
                worksheet.Cell(row, 7).Value = tetel.TetelNev;
                worksheet.Cell(row, 9).Value = tetel.HallgatoFelvetele;
                worksheet.Cell(row, 10).Value = tetel.Szak;
                worksheet.Cell(row, 11).Value = tetel.PenzugyiStatusz;

                row++;
            }

            workbook.SaveAs($"tablak/oe_neptunba_NIK_{Config.AktualisHonap}_koz_nyers.xlsx");
        }

        static void Osszesito()
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Munka1");

            worksheet.Cell(1, 1).Value = "Neptun kód";
            worksheet.Cell(1, 2).Value = "Képzéskód";
            worksheet.Cell(1, 3).Value = "Jogcím";
            worksheet.Cell(1, 4).Value = "Indoklás";
            worksheet.Cell(1, 5).Value = "Kiutalandó összeg";

            int row = 2;
            foreach (Tetel tetel in tetelek.Where(x => x.Jogcim.Contains("Rendszeres")))
            {
                worksheet.Cell(row, 1).Value = tetel.Owner;
                worksheet.Cell(row, 2).Value = tetel.Kepzes;
                worksheet.Cell(row, 3).Value = tetel.Jogcim;
                worksheet.Cell(row, 4).Value = tetel.Indoklas;
                worksheet.Cell(row, 5).Value = tetel.Value;

                row++;
            }
            foreach (Tetel tetel in tetelek.Where(x => x.Jogcim.Contains("Egyszeri")).OrderBy(x => x.Owner))
            {
                worksheet.Cell(row, 1).Value = tetel.Owner;
                worksheet.Cell(row, 2).Value = tetel.Kepzes;
                worksheet.Cell(row, 3).Value = tetel.Jogcim;
                worksheet.Cell(row, 4).Value = tetel.Indoklas;
                worksheet.Cell(row, 5).Value = tetel.Value;

                row++;
            }
            workbook.SaveAs($"tablak/osszesito_{Config.AktualisHonap}.xlsx");
        }
    }
}
