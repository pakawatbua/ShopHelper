using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShopHelper
{
    class Program_O
    {
        //private const string RootPath = @"C:\Users\pbuaklay\Google Drive\Shop";

        //static void Main()
        //{
        //    Console.WriteLine("Enter fucntion : ");
        //    Console.WriteLine("Profit calculation 'Lazada' press 1");
        //    Console.WriteLine("Profit calculation 'Shopee' press 2");

        //    var fucntion = (Common.Function)Console.Read();

        //    switch (fucntion)
        //    {
        //        case Common.Function.ProfitcalLaz:
        //            Run(() => CalculateProfitLaz(
        //                    new PathCombination(
        //                Path.Combine(RootPath, @"Keep\Cost.xlsx"),
        //                Path.Combine(RootPath, @"Functions\ProfitcalLaz\sell.xlsx"),
        //                Path.Combine(RootPath, $@"Functions\ProfitcalLaz\profit_{new Random().Next()}_{DateTime.Now.Date.Day}.xlsx"))));

        //            break;
        //        case Common.Function.ProfitcalSho:
        //            CalculateProfitSho(
        //                Path.Combine(RootPath, @"Keep\Cost.xlsx"),
        //                Path.Combine(RootPath, @"Functions\ProfitcalSho\sell.xlsx"),
        //                Path.Combine(RootPath, $@"Functions\ProfitcalSho\profit_{new Random().Next()}_{DateTime.Now.Date.Day}.xlsx"));
        //            break;
        //        default:
        //            break;
        //    }

        //    //ComparePrice(0.3);
        //    //ProfitCal();


        //    //var costPath = Path.Combine(RootPath, @"Price\Shopee\price.xlsx");
        //    //var sellPath = Path.Combine(RootPath, @"Price\Lazada\price.xlsx");
        //    //var profitPath = Path.Combine(RootPath, $@"Price\Result\profit_{DateTime.Now.Date.Day}.xlsx");
        //    //ComparePrice(Common.Shop.Shopee, Common.Shop.Lazada, costPath, sellPath, profitPath);

        //    //var sourcePath = Path.Combine(RootPath, @"Profit\Shopee\cost.xlsx");
        //    //var descPath = Path.Combine(RootPath, @"Profit\Lazada\sell.xlsx");
        //    //var profitPath = Path.Combine(RootPath, $@"Profit\Profit\profit_{DateTime.Now.Millisecond}.xlsx");
        //    //CalculateProfit(Common.Shop.Shopee, Common.Shop.Lazada, sourcePath, descPath, profitPath);

        //    //var scrPath = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
        //    //var descPath = Path.Combine(RootPath, @"Stock\Pun\stock.xlsx");
        //    //var outputPath = Path.Combine(RootPath, $@"Stock\Shopee\updated_X_{DateTime.Now.Day}.xlsx");
        //    //UpdateStock(Common.Shop.Shopee, Common.Shop.Shopee, scrPath, descPath, outputPath);

        //    //var scrPath = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
        //    //var descPath = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
        //    //var outputPath = Path.Combine(RootPath, $@"Stock\Lazada\Update\updated{DateTime.Now.Day}.xlsx");
        //    //UpdateStock(Common.Shop.Lazada, Common.Shop.Shopee, scrPath, descPath, outputPath);

        //    //var scrPath = Path.Combine(RootPath, @"Stock\Lazada\stock.xlsx");
        //    //var descPath = Path.Combine(RootPath, @"Stock\Shopee\stock.xlsx");
        //    //var outputPath = Path.Combine(RootPath, $@"Stock\Lazada\UnmatchedName\unmatchedName{DateTime.Now.Day}.xlsx");
        //    //GetUnmatchedName(Common.Shop.Lazada, Common.Shop.Shopee, scrPath, descPath, outputPath);

        //    //var top100Path = Path.Combine(RootPath, @"Price\Lazada\Top100.xlsx");
        //    //var allPath = Path.Combine(RootPath, @"Price\Lazada\all.xlsx");
        //    //var outputPath = Path.Combine(RootPath, $@"Price\Lazada\Top100_{DateTime.Now.Day}.xlsx");
        //    //GetTop100Price(top100Path, allPath, outputPath);


        //    //var scrPath = Path.Combine(RootPath, @"Price\Shopee\price.xlsx");
        //    //var descPath = Path.Combine(RootPath, @"Price\Pun\price.xlsx");
        //    //var outputPath = Path.Combine(RootPath, $@"Price\Shopee\updated{DateTime.Now.Day}.xlsx");
        //    //UpdatePrice(Common.Shop.Shopee, Common.Shop.Shopee, scrPath, descPath, outputPath);
        //}

        //private static void Run(Action function)
        //{
        //    try
        //    {
        //        function();

        //        Console.WriteLine("Done!!!");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error :" + ex.Message);
        //    }

        //    Console.ReadKey();
        //}

        //private static void CalculateProfitLaz(PathCombination paths)
        //{
        //    var cost = new File(Common.Shop.Shopee, Common.Type.Cost).Read(paths.FirstPath);
        //    var sell = new File(Common.Shop.Lazada, Common.Type.Sell).Read(paths.SecoundPath);

        //    new ProfitCalculator(cost, sell).Write(Common.Shop.Lazada, Path.Combine(paths.OutPutPath));
        //}

        ////private static void GetTop100Price(string top100Path, string allPath, string outputPath)
        ////{
        ////    try
        ////    {
        ////        var top100Products = new File(Common.Shop.Lazada, Common.Type.Product).Read(top100Path);

        ////        new CampaignManager(top100Products).Write(Common.Shop.Lazada,
        ////            Path.Combine(outputPath) , allPath);

        ////        Console.WriteLine("Done!!!");
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine("Error :" + ex.Message);
        ////    }

        ////    Console.ReadKey();
        ////}

        ////private static void GetUnmatchedName(Common.Shop srcShop, Common.Shop descShop, string scrPath, string descPath,
        ////    string outputPath)
        ////{
        ////    try
        ////    {
        ////        var sourceStock = new File(srcShop, Common.Type.Stock).Read(scrPath);

        ////        var descStock = new File(descShop, Common.Type.Stock).Read(descPath);

        ////        new UnmatchedNameManager(sourceStock, descStock).Write(srcShop,
        ////            Path.Combine(outputPath));

        ////        Console.WriteLine("Done!!!");
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine("Error :" + ex.Message);
        ////    }

        ////    Console.ReadKey();
        ////}

        ////private static void UpdatePrice(Common.Shop srcShop, Common.Shop descShop, string scrPath, string descPath, string outputPath)
        ////{
        ////    try
        ////    {
        ////        var sourceStock = new File(srcShop, Common.Type.Price).Read(scrPath);

        ////        var descStock = new File(descShop, Common.Type.Price).Read(descPath);

        ////        new PriceManager(sourceStock, descStock).Write(srcShop,
        ////            Path.Combine(outputPath));

        ////        Console.WriteLine("Done!!!");
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine("Error :" + ex.Message);
        ////    }

        ////    Console.ReadKey();
        ////}

        ////private static void UpdateStock(Common.Shop srcShop, Common.Shop descShop, string scrPath, string descPath, string outputPath)
        ////{
        ////    try
        ////    {
        ////        var sourceStock = new File(srcShop, Common.Type.Stock).Read(scrPath);

        ////        var descStock = new File(descShop, Common.Type.Stock).Read(descPath);

        ////        new StockManager(sourceStock, descStock).Write(srcShop,
        ////            Path.Combine(outputPath));

        ////        Console.WriteLine("Done!!!");
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine("Error :" + ex.Message);
        ////    }

        ////    Console.ReadKey();
        ////}

        ////private static void ProfitCal()
        ////{
        ////    try
        ////    {
        ////        var lazadaSellList = new LazadaSellListFile().Read(Path.Combine(RootPath, "lazadaSellList.xlsx"));

        ////        var lazadaBasePrice = new LazadaBasePriceFile().Read(Path.Combine(RootPath, "lazadaBasePrice.xlsx"));

        ////        new ProfitFile(lazadaSellList, lazadaBasePrice).Write(Path.Combine(RootPath, $"profited_{ DateTime.Now}.xlsx"));

        ////        Console.WriteLine("Done!!!");
        ////    }
        ////    catch (Exception ex)
        ////    {

        ////        Console.WriteLine("Error :" + ex.Message);
        ////    }

        ////    Console.ReadKey();
        ////}

        ////private static void ComparePrice(Common.Shop sourceShop, Common.Shop descShop, string sourcePath, string descPath, string profitPath)
        ////{
        ////    try
        ////    {
        ////        var shopeePriceList = new File(sourceShop, Common.Type.Price).Read(sourcePath);

        ////        var lazadaPriceList = new File(descShop, Common.Type.Price).Read(descPath);

        ////        new PriceManager(shopeePriceList, lazadaPriceList).Write(descShop,
        ////            Path.Combine(profitPath));

        ////        Console.WriteLine("Done!!!");
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine("Error :" + ex.Message);
        ////    }

        ////    Console.ReadKey();
        ////}
    }
}
