using System;
using System.Collections.Generic;
using System.Linq;

namespace ShopHelper
{
    internal class PriceComparer
    {
        public PriceComparer()
        {
        }

        public IEnumerable<ComparedItem> Compare(IEnumerable<Item> lazadaItems, IEnumerable<Item> shopeeItems, double torerantRate)
        {
            var shopee = shopeeItems.ToList();
            foreach (var l in lazadaItems.ToList())
            {
                var acceptedCount = l.Name.Length * torerantRate;
                var matchedShopee = shopee.FirstOrDefault(s => Compute(s.Name.ToLower(), l.Name.ToLower()) < acceptedCount);
                //var comparedItems = new ComparedItem() { LazadaName = l.Name, LazadaPrice = l.Price };

                //if (matchedShopee != null)
                //{
                //    comparedItems.ShopeeName = matchedShopee.Name;
                //    comparedItems.ShopeePrice = matchedShopee.Price;
                //}

                //yield return comparedItems;

                

                if (matchedShopee != null)
                {
                    var comparedItems = new ComparedItem() { LazadaName = l.Name, LazadaPrice = l.Price };
                    comparedItems.ShopeeName = matchedShopee.Name;
                    comparedItems.ShopeePrice = matchedShopee.Price;
                    yield return comparedItems;
                }

                
            }
        }

        internal object Write(string path)
        {
            throw new NotImplementedException();
        }

        private int Compute(string s, string t)
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