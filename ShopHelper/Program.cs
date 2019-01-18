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
            //var rootPath = @"C:\Shop";

            //var lazadaItems =  new LazadaFile().Read(Path.Combine(rootPath, "lazada.xlsx"));

            //var shopeeItems = new ShopeeFile().Read(Path.Combine(rootPath, "shopee.xlsx"));

            //var comparedItems = new PriceComparer().Compare(lazadaItems, shopeeItems);

            //var isSuccess = new ResultFile(comparedItems).Write(Path.Combine(rootPath, "compared.xlsx"));

            Console.WriteLine(Compute("Dr.andrew weil for origins mega-mushroom relief & resilience 50 ml (สูตรใหม่)", "Dr.andrew weil for origins mega-mushroom relief & resilience , 50 ml, 100ml (สูตรใหม่)"));
            Console.WriteLine(Compute("yves saint laurent Couture Blush", "yves saint laurent Couture Blush"));
            Console.WriteLine(Compute("Kiehl’s powerful strength line reducing 12.5 vitamin c 50Ml", "Kiehl’s powerful strength line reducing 12.5 vitamin c 50Ml , 75ml"));
            Console.ReadKey();
        }

        private static int Compute(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // Step 1
            if (n == 0)
            {
                return m;
            }

            if (m == 0)
            {
                return n;
            }

            // Step 2
            for (int i = 0; i <= n; d[i, 0] = i++)
            {
            }

            for (int j = 0; j <= m; d[0, j] = j++)
            {
            }

            // Step 3
            for (int i = 1; i <= n; i++)
            {
                //Step 4
                for (int j = 1; j <= m; j++)
                {
                    // Step 5
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;

                    // Step 6
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            // Step 7
            return d[n, m];
        }
    }
}
