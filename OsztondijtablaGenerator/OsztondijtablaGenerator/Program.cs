using System;

namespace OsztondijtablaGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            NeptunAdatok.Beolvas();
            RendszeresOsztondij.Beolvas();
            EgyszeriOsztondij.Beolvas();
            Tetel.TetelGeneralas();
        }
    }
}
