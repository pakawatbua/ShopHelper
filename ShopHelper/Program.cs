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
            UpdateStock(Common.Shop.Shopee, Common.Shop.Shopee);
        }

        private static void UpdateStock(Common.Shop sourceShop, Common.Shop descShop)
        {
            try
            {
                var sourceStock = new File(sourceShop, Common.Type.Stock).Read(Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx"));

                var descStock = new File(descShop, Common.Type.Stock).Read(Path.Combine(RootPath, @"Stock\Pun\stock.xlsx"));

                new StockManager(sourceStock, descStock).Write(Common.Shop.Shopee,
                    Path.Combine(RootPath, $@"Stock\Shopee\updated{DateTime.Now.ToShortDateString()}.xlsx"));

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
