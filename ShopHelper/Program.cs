using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShopHelper
{
    class Program
    {
        private const string _rootPath = @"C:\Shop";

        static void Main(string[] args)
        {
            Compare(0.3);
        }

        static void Compare(double torerantRate)
        {
            try
            {
                var lazadaItems = new LazadaFile().Read(Path.Combine(_rootPath, "lazada.xlsx"));

                var shopeeItems = new ShopeeFile().Read(Path.Combine(_rootPath, "shopee.xlsx"));

                var comparedItems = new PriceComparer().Compare(lazadaItems, shopeeItems, torerantRate);

                var isSuccess = new ComparedFile(comparedItems).Write(Path.Combine(_rootPath, $"compared_{ DateTime.Now.ToString("yyyyMMddTHHmmss") }.xlsx"));

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
