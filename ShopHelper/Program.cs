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
        static void Main(string[] args)
        {
            var rootPath = @"C:\Shop";

            var lazadaItems =  new LazadaFile().Read(Path.Combine(rootPath, "lazada.xlsx"));

            var shopeeItems = new ShopeeFile().Read(Path.Combine(rootPath, "shopee.xlsx"));

            var comparedItems = new PriceComparer().Compare(lazadaItems, shopeeItems);

            var isSuccess = new ResultFile(comparedItems).Write(Path.Combine(rootPath, "compared.xlsx"));
        }
    }
}
