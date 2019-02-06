using System;
using System.IO;

namespace ShopHelper
{
    internal class Program
    {
        private const string RootPath = @"C:\Shop";

        static void Main(string[] args)
        {
            ComparePrice(0.3);
        }

        static void ComparePrice(double torerantRate)
        {
            try
            {
                var lazadaPrice = new LazadaFile().Read(Path.Combine(RootPath, "lazada.xlsx"));

                var shopeePrice = new ShopeeFile().Read(Path.Combine(RootPath, "shopee.xlsx"));

                var comparedPrice = new PriceComparer().Compare(lazadaPrice, shopeePrice, torerantRate);

                var isSuccess = new ComparedFile(comparedPrice).Write(Path.Combine(RootPath, $"compared_{ DateTime.Now :yyyyMMddTHHmmss}.xlsx"));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {

                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        static void ProfitCal()
        {
            try
            {
                var lazadaSellList = new LazadaSellListFile().Read(Path.Combine(RootPath, "lazadaSellList.xlsx"));

                var lazadaBasePrice = new LazadaBasePriceFile().Read(Path.Combine(RootPath, "lazadaBasePrice.xlsx"));

                var isSuccess = new ProfitFile(lazadaSellList, lazadaBasePrice).Write(Path.Combine(RootPath, $"profited_{ DateTime.Now :yyyyMMddTHHmmss}.xlsx"));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {

                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }
    }
}
