using System;
using System.Collections.Generic;
using System.IO;

namespace ShopHelper
{
    internal class Program
    {
        private const string RootPath = @"C:\Shop";

        static void Main()
        {
            //ComparePrice(0.3);
            //ProfitCal();

            //var scrPart = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var descPart = Path.Combine(RootPath, @"Stock\Pun\stock.xlsx");
            //var outputPart = Path.Combine(RootPath, $@"Stock\Shopee\updated{DateTime.Now.ToShortDateString()}.xlsx");
            //UpdateStock(Common.Shop.Shopee, Common.Shop.Shopee, scrPart, descPart, outputPart);

            var scrPart = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
            var descPart = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            var outputPart = Path.Combine(RootPath, $@"Stock\Lazada\Update\updated{DateTime.Now.Day}.xlsx");
            UpdateStock(Common.Shop.Lazada, Common.Shop.Shopee, scrPart, descPart, outputPart);

            //var scrPart = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
            //var descPart = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var outputPart = Path.Combine(RootPath, $@"Stock\Lazada\UnmatchedName\unmatchedName{DateTime.Now.Day}.xlsx");
            //GetUnmatchedName(Common.Shop.Lazada, Common.Shop.Shopee, scrPart, descPart, outputPart);
        }

        private static void GetUnmatchedName(Common.Shop srcShop, Common.Shop descShop, string scrPart, string descPart,
            string outputPart)
        {
            try
            {
                var sourceStock = new File(srcShop, Common.Type.Stock).Read(scrPart);

                var descStock = new File(descShop, Common.Type.Stock).Read(descPart);

                new UnmatchedNameManager(sourceStock, descStock).Write(srcShop,
                    Path.Combine(outputPart));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void UpdateStock(Common.Shop srcShop, Common.Shop descShop, string scrPart, string descPart, string outputPart)
        {
            try
            {
                var sourceStock = new File(srcShop, Common.Type.Stock).Read(scrPart);

                var descStock = new File(descShop, Common.Type.Stock).Read(descPart);

                new StockManager(sourceStock, descStock).Write(srcShop,
                    Path.Combine(outputPart));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void ComparePrice(double torerantRate)
        {
            try
            {
                var lazadaPrice = new LazadaFile().Read(Path.Combine(RootPath, "lazada.xlsx"));

                var shopeePrice = new ShopeeFile().Read(Path.Combine(RootPath, "shopee.xlsx"));

                new ComparedFile(lazadaPrice, shopeePrice, torerantRate).Write(Path.Combine(RootPath, $"compared_{ DateTime.Now:yyyyMMddTHHmmss}.xlsx"));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {

                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void ProfitCal()
        {
            try
            {
                var lazadaSellList = new LazadaSellListFile().Read(Path.Combine(RootPath, "lazadaSellList.xlsx"));

                var lazadaBasePrice = new LazadaBasePriceFile().Read(Path.Combine(RootPath, "lazadaBasePrice.xlsx"));

                new ProfitFile(lazadaSellList, lazadaBasePrice).Write(Path.Combine(RootPath, $"profited_{ DateTime.Now:yyyyMMddTHHmmss}.xlsx"));

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
