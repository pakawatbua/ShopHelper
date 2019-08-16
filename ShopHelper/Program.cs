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

            var costPath = Path.Combine(RootPath, @"Profit\Shopee\cost.xlsx");
            var sellPath = Path.Combine(RootPath, @"Profit\Lazada\sell.xlsx");
            var profitPath = Path.Combine(RootPath, $@"Profit\Profit\profit_{DateTime.Now.Millisecond}.xlsx");
            CalculateProfit(Common.Shop.Shopee, Common.Shop.Lazada, costPath, sellPath, profitPath);

            //var scrPath = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var descPath = Path.Combine(RootPath, @"Stock\Pun\stock.xlsx");
            //var outputPath = Path.Combine(RootPath, $@"Stock\Shopee\updated_X_{DateTime.Now.Day}.xlsx");
            //UpdateStock(Common.Shop.Shopee, Common.Shop.Shopee, scrPath, descPath, outputPath);

            //var scrPath = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
            //var descPath = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var outputPath = Path.Combine(RootPath, $@"Stock\Lazada\Update\updated{DateTime.Now.Day}.xlsx");
            //UpdateStock(Common.Shop.Lazada, Common.Shop.Shopee, scrPath, descPath, outputPath);

            //var scrPath = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
            //var descPath = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
            //var outputPath = Path.Combine(RootPath, $@"Stock\Lazada\UnmatchedName\unmatchedName{DateTime.Now.Day}.xlsx");
            //GetUnmatchedName(Common.Shop.Lazada, Common.Shop.Shopee, scrPath, descPath, outputPath);

            //var top100Path = Path.Combine(RootPath, @"Price\Lazada\Top100.xlsx");
            //var allPath = Path.Combine(RootPath, @"Price\Lazada\all.xlsx");
            //var outputPath = Path.Combine(RootPath, $@"Price\Lazada\Top100_{DateTime.Now.Day}.xlsx");
            //GetTop100Price(top100Path, allPath, outputPath);


            //var scrPath = Path.Combine(RootPath, @"Price\Shopee\price.xlsx");
            //var descPath = Path.Combine(RootPath, @"Price\Pun\price.xlsx");
            //var outputPath = Path.Combine(RootPath, $@"Price\Shopee\updated{DateTime.Now.Day}.xlsx");
            //UpdatePrice(Common.Shop.Shopee, Common.Shop.Shopee, scrPath, descPath, outputPath);
        }

        private static void GetTop100Price(string top100Path, string allPath, string outputPath)
        {
            try
            {
                var top100Products = new File(Common.Shop.Lazada, Common.Type.Product).Read(top100Path);

                new CampaignManager(top100Products).Write(Common.Shop.Lazada,
                    Path.Combine(outputPath) , allPath);

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void GetUnmatchedName(Common.Shop srcShop, Common.Shop descShop, string scrPath, string descPath,
            string outputPath)
        {
            try
            {
                var sourceStock = new File(srcShop, Common.Type.Stock).Read(scrPath);

                var descStock = new File(descShop, Common.Type.Stock).Read(descPath);

                new UnmatchedNameManager(sourceStock, descStock).Write(srcShop,
                    Path.Combine(outputPath));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void UpdatePrice(Common.Shop srcShop, Common.Shop descShop, string scrPath, string descPath, string outputPath)
        {
            try
            {
                var sourceStock = new File(srcShop, Common.Type.Price).Read(scrPath);

                var descStock = new File(descShop, Common.Type.Price).Read(descPath);

                new PriceManager(sourceStock, descStock).Write(srcShop,
                    Path.Combine(outputPath));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void CalculateProfit(Common.Shop costShop, Common.Shop sellShop, string costPath, string sellPath, string profitPath)
        {
            try
            {
                var cost = new File(costShop, Common.Type.Price).Read(costPath);

                var sell = new File(sellShop, Common.Type.Sell).Read(sellPath);

                new ProfitCalculator(cost, sell).Write(sellShop,
                    Path.Combine(profitPath));

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void UpdateStock(Common.Shop srcShop, Common.Shop descShop, string scrPath, string descPath, string outputPath)
        {
            try
            {
                var sourceStock = new File(srcShop, Common.Type.Stock).Read(scrPath);

                var descStock = new File(descShop, Common.Type.Stock).Read(descPath);

                new StockManager(sourceStock, descStock).Write(srcShop,
                    Path.Combine(outputPath));

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
