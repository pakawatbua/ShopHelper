using ShopHelper.Commons;
using ShopHelper.Models;
using System;
using System.IO;

namespace ShopHelper
{
    internal class Program
    {
        private const string RootPath = @"C:\Users\pbuaklay\Google Drive\Shop";

        static void Main()
        {
            Console.WriteLine("Enter fucntion : ");
            Console.WriteLine("Profit calculation 'Lazada' press 0");
            Console.WriteLine("Profit calculation 'Shopee' press 1");
            Console.WriteLine("Update stock 'BYM' press 2");
            Console.WriteLine("Update stock 'Laz' press 3");

            Common.Function fucntion;
            Enum.TryParse(Console.ReadLine(), out fucntion);

            Console.WriteLine("");

            switch (fucntion)
            {
                case Common.Function.ProfitcalLaz:
                    Run(() => CalculateProfitLaz(
                            new PathCombination(
                        Path.Combine(RootPath, @"Keep\Cost.xlsx"),
                        Path.Combine(RootPath, @"Functions\ProfitcalLaz\sell.xlsx"),
                        Path.Combine(RootPath, $@"Functions\ProfitcalLaz\profit_{new Random().Next()}_{DateTime.Now.Date.Month}_{DateTime.Now.Date.Day}.xlsx"))));

                    break;
                case Common.Function.ProfitcalSho:
                    Run(() => CalculateProfitSho(
                            new PathCombination(
                        Path.Combine(RootPath, @"Keep\Cost.xlsx"),
                        Path.Combine(RootPath, @"Functions\ProfitcalSho\sell_8.xlsx"),
                        Path.Combine(RootPath, $@"Functions\ProfitcalSho\profit_{new Random().Next()}_{DateTime.Now.Date.Month}_{DateTime.Now.Date.Day}.xlsx"))));

                    break;
                case Common.Function.UpdateStockBYM:
                    Run(() => UpdateStockBYM(
                            new PathCombination(
                        Path.Combine(RootPath, @"Functions\UpdateStockBYM\shoStock.xlsx"),
                        Path.Combine(RootPath, @"Functions\UpdateStockBYM\bymStock.xlsx"),
                        Path.Combine(RootPath, $@"Functions\UpdateStockBYM\updatedBYMStock_{new Random().Next()}_{DateTime.Now.Date.Month}_{DateTime.Now.Date.Day}.xlsx"))));

                    break;
                case Common.Function.UpdateStockLaz:
                    Run(() => UpdateStockLaz(
                            new PathCombination(
                        Path.Combine(RootPath, @"Functions\UpdateStockLaz\shoStock.xlsx"),
                        Path.Combine(RootPath, @"Functions\UpdateStockLaz\lazStock_test.xlsx"),
                        Path.Combine(RootPath, $@"Functions\UpdateStockLaz\updatedLazStock_{new Random().Next()}_{DateTime.Now.Date.Month}_{DateTime.Now.Date.Day}.xlsx"))));

                    break;
                default:
                    break;
            }
        }

        private static void Run(Action function)
        {
            try
            {
                function();

                Console.WriteLine("Done!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }

            Console.ReadKey();
        }

        private static void CalculateProfitLaz(PathCombination paths)
        {
            Console.WriteLine("Calculating profit of lazada...");

            var cost = new File(Common.Shop.Shopee, Common.Type.Cost).Read(paths.FirstPath);
            var sell = new File(Common.Shop.Lazada, Common.Type.Sell).Read(paths.SecoundPath);

            new ProfitCalculator(cost, sell).Write(Common.Shop.Lazada, Path.Combine(paths.OutPutPath));
        }

        private static void CalculateProfitSho(PathCombination paths)
        {
            Console.WriteLine("Calculating profit of shopee...");

            var cost = new File(Common.Shop.Shopee, Common.Type.Cost).Read(paths.FirstPath);
            var sell = new File(Common.Shop.Shopee, Common.Type.Sell).Read(paths.SecoundPath);

            new ProfitCalculator(cost, sell).Write(Common.Shop.Shopee, Path.Combine(paths.OutPutPath));
        }

        private static void UpdateStockBYM(PathCombination paths)
        {
            Console.WriteLine("Updateiong stock of BYM...");

            var baseStock = new File(Common.Shop.Shopee, Common.Type.Stock).Read(paths.FirstPath);
            var targetStock = new File(Common.Shop.BYM, Common.Type.Stock).Read(paths.SecoundPath);

            new StockUpdater(baseStock, targetStock).Write(Common.Shop.BYM, Path.Combine(paths.OutPutPath));
        }

        private static void UpdateStockLaz(PathCombination paths)
        {
            Console.WriteLine("Updateiong stock of Laz...");

            var baseStock = new File(Common.Shop.Shopee, Common.Type.Stock).Read(paths.FirstPath);
            var targetStock = new File(Common.Shop.Lazada, Common.Type.Stock).Read(paths.SecoundPath);

            new StockUpdater(baseStock, targetStock).Write(Common.Shop.Lazada, Path.Combine(paths.OutPutPath));
        }
    }
}
