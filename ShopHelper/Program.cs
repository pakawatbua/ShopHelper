using System;
using System.Collections.Generic;
using System.IO;

namespace ShopHelper
{
    internal class Program
    {
        private const string RootPath = @"C:\Users\pbuaklay\OneDrive\Shop";

        static void Main()
        {
            //ComparePrice(0.3);
            //ProfitCal();

            var scrPart = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            var descPart = Path.Combine(RootPath, @"Stock\Pun\stock.xlsx");
            var outputPart = Path.Combine(RootPath, $@"Stock\Shopee\updated_{DateTime.Now.Day}.xlsx");
            UpdateStock(Common.Shop.Shopee, Common.Shop.Shopee, scrPart, descPart, outputPart);

            //var scrPart = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
            //var descPart = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var outputPart = Path.Combine(RootPath, $@"Stock\Lazada\Update\updated{DateTime.Now.Day}.xlsx");
            //UpdateStock(Common.Shop.Lazada, Common.Shop.Shopee, scrPart, descPart, outputPart);

            //var scrPart = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
            //var descPart = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var outputPart = Path.Combine(RootPath, $@"Stock\Lazada\UnmatchedName\unmatchedName{DateTime.Now.Day}.xlsx");
            //GetUnmatchedName(Common.Shop.Lazada, Common.Shop.Shopee, scrPart, descPart, outputPart);

            //var top100Part = Path.Combine(RootPath, @"Price\Lazada\Top100.xlsx");
            //var allPart = Path.Combine(RootPath, @"Price\Lazada\all.xlsx");
            //var outputPart = Path.Combine(RootPath, $@"Price\Lazada\Top100_{DateTime.Now.Day}.xlsx");
            //GetTop100Price(top100Part, allPart, outputPart);


            //var scrPart = Path.Combine(RootPath, @"Price\Shopee\price.xlsx");
            //var descPart = Path.Combine(RootPath, @"Price\Pun\price.xlsx");
            //var outputPart = Path.Combine(RootPath, $@"Price\Shopee\updated{DateTime.Now.Day}.xlsx");
            //UpdatePrice(Common.Shop.Shopee, Common.Shop.Shopee, scrPart, descPart, outputPart);
        }

        private static void GetTop100Price(string top100Part, string allPart, string outputPart)
        {
            try
            {
                var top100Products = new File(Common.Shop.Lazada, Common.Type.Product).Read(top100Part);

                new CampaignManager(top100Products).Write(Common.Shop.Lazada,
                    Path.Combine(outputPart) , allPart);

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
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

        private static void UpdatePrice(Common.Shop srcShop, Common.Shop descShop, string scrPart, string descPart, string outputPart)
        {
            try
            {
                var sourceStock = new File(srcShop, Common.Type.Price).Read(scrPart);

                var descStock = new File(descShop, Common.Type.Price).Read(descPart);

                new PriceManager(sourceStock, descStock).Write(srcShop,
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
