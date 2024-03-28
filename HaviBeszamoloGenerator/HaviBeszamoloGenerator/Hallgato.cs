using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HaviBeszamoloGenerator
{
    class Hallgato
    {
        public static List<Hallgato> Hallgatok = new List<Hallgato>();

        public string Neptun { get; set; }
        public string KepzesKod { get; set; }
        public int Tanulmanyi { get; set; }
        public int KozeletiRendszeres { get; set; }
        public int KozeletiEgyszeri { get; set; }

        public int Osszeg
        {
            get
            {
                return this.Tanulmanyi + this.KozeletiRendszeres + this.KozeletiEgyszeri;
            }
        }

        public Hallgato(string neptun, string kepzesKod, int tanulmanyi, int kozeletiRendszeres, int kozeletiEgyszeri)
        {
            this.Neptun = neptun;
            this.KepzesKod = kepzesKod;
            this.Tanulmanyi = tanulmanyi;
            this.KozeletiRendszeres = kozeletiRendszeres;
            this.KozeletiEgyszeri = kozeletiEgyszeri;
        }

        public Hallgato(string neptun, string kepzesKod, int tanulmanyi)
        {
            this.Neptun = neptun;
            this.KepzesKod = kepzesKod;
            this.Tanulmanyi = tanulmanyi;
            this.KozeletiRendszeres = 0;
            this.KozeletiEgyszeri = 0;
        }

        public override bool Equals(object? obj)
        {
            if (obj is null) return false;
            if (obj is not Hallgato) return false;

            return (obj as Hallgato).Neptun.Equals(this.Neptun);
        }

        public static void TanulmanyiBeolvas()
        {
            StreamWriter sw = new StreamWriter("logs/tanulmanyi.log");

            var tanulmanyiWorksheet = new XLWorkbook(Config.tanulmanyiFajl).Worksheet(1);

            if (tanulmanyiWorksheet.IsEmpty()) return;

            for (int i = 2; i <= tanulmanyiWorksheet.LastRowUsed().RowNumber(); i++)
            {
                var row = tanulmanyiWorksheet.Row(i);

                string neptun = row.Cell(1).GetString().Trim();
                string kepzeskod = row.Cell(2).GetString().Trim();
                int tanulmanyi = int.Parse(row.Cell(4).GetString().Trim());

                Hallgatok.Add(
                    new Hallgato(
                        neptun,
                        kepzeskod,
                        tanulmanyi
                ));

                sw.WriteLine($"{neptun} - {tanulmanyi}");
            }

            sw.Close();
        }

        public static void KozeletiBeolvas()
        {
            StreamWriter sw = new StreamWriter("logs/kozeleti.log");

            var kozeletiWorksheet = new XLWorkbook(Config.kozeletiFajl).Worksheet(1);

            for (int i = 2; i <= kozeletiWorksheet.LastRowUsed().RowNumber(); i++)
            {
                var row = kozeletiWorksheet.Row(i);

                Hallgato hallgato;

                if (row.Cell(7).GetString().Contains("Rendszeres"))
                {
                    string neptun = row.Cell(1).GetString().Trim();
                    string kepzeskod = row.Cell(2).GetString().Trim();
                    int rendszeres = int.Parse(row.Cell(4).GetString().Trim());

                    hallgato = new Hallgato(
                        neptun,
                        kepzeskod,
                        0,
                        rendszeres,
                        0
                    );

                    if (Hallgatok.Contains(hallgato))
                    {
                        int index = Hallgatok.IndexOf(hallgato);
                        int elozo_rendszeres = Hallgatok[index].KozeletiEgyszeri; 


                        Hallgatok[index].KozeletiRendszeres += hallgato.KozeletiRendszeres;

                        sw.WriteLine($"{Hallgatok[index].Neptun} - Rendszeres ({elozo_rendszeres} --> {Hallgatok[index].KozeletiRendszeres})");
                    }
                    else
                    {
                        Hallgatok.Add(hallgato);

                        sw.WriteLine($"{hallgato.Neptun} - Rendszeres (0 --> {hallgato.KozeletiRendszeres})");
                    }
                }

                if (row.Cell(7).GetString().Contains("Egyszeri"))
                {
                    string neptun = row.Cell(1).GetString().Trim();
                    string kepzeskod = row.Cell(2).GetString().Trim();
                    int egyszeri = int.Parse(row.Cell(4).GetString().Trim());

                    hallgato = new Hallgato(
                        neptun,
                        kepzeskod,
                        0,
                        0,
                        egyszeri
                    );

                    if (Hallgatok.Contains(hallgato))
                    {
                        int index = Hallgatok.IndexOf(hallgato);
                        int elozo_egyszeri = Hallgatok[index].KozeletiEgyszeri;

                        Hallgatok[index].KozeletiEgyszeri += hallgato.KozeletiEgyszeri;

                        sw.WriteLine($"{Hallgatok[index].Neptun} - Egyszeri ({elozo_egyszeri} --> {Hallgatok[index].KozeletiEgyszeri})");
                    }
                    else
                    {
                        Hallgatok.Add(hallgato);

                        sw.WriteLine($"{hallgato.Neptun} - Egyszeri (0 --> {hallgato.KozeletiRendszeres})");
                    }
                }
            }

            sw.Close();
        }

        public static void TablaGeneral()
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Munka1");
            worksheet.Cell(1, 1).Value = "Neptun";
            worksheet.Cell(1, 2).Value = "Évf";
            worksheet.Cell(1, 3).Value = "Tank";
            worksheet.Cell(1, 4).Value = "Képzés kód";
            worksheet.Cell(1, 5).Value = "Összeg";
            worksheet.Cell(1, 6).Value = "Tan.R.";
            worksheet.Cell(1, 7).Value = "Tan.1X";
            worksheet.Cell(1, 8).Value = "KHÖK.R";
            worksheet.Cell(1, 9).Value = "KHÖK.1x";

            int row = 2;

            foreach (Hallgato hallgato in Hallgatok.OrderBy(x => x.Neptun))
            {
                worksheet.Cell(row, 1).Value = hallgato.Neptun;
                worksheet.Cell(row, 2).Value = "";
                worksheet.Cell(row, 3).Value = "";
                worksheet.Cell(row, 4).Value = hallgato.KepzesKod;
                worksheet.Cell(row, 5).Value = hallgato.Osszeg;
                worksheet.Cell(row, 6).Value = hallgato.Tanulmanyi;
                worksheet.Cell(row, 7).Value = 0;
                worksheet.Cell(row, 8).Value = hallgato.KozeletiRendszeres;
                worksheet.Cell(row, 9).Value = hallgato.KozeletiEgyszeri;

                row++;
            }
            workbook.SaveAs($"tablak/beszamolo_{Config.AktualisHonap}.xlsx");
        }
    }
}
