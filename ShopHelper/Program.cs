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
                        Path.Combine(RootPath, $@"Functions\ProfitcalLaz\profit_{new Random().Next()}_{DateTime.Now.Date.Day}.xlsx"))));

                    break;
                case Common.Function.ProfitcalSho:
                    Run(() => CalculateProfitSho(
                            new PathCombination(
                        Path.Combine(RootPath, @"Keep\Cost.xlsx"),
                        Path.Combine(RootPath, @"Functions\ProfitcalSho\sell.xlsx"),
                        Path.Combine(RootPath, $@"Functions\ProfitcalSho\profit_{new Random().Next()}_{DateTime.Now.Date.Day}.xlsx"))));

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
    }
}
