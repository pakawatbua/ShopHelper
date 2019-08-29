using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using System;
using System.Globalization;
using ShopHelper.Commons;
using ShopHelper.Models;


namespace ShopHelper
{
    public class ComparedFile
    {
        private readonly IEnumerable<Item> _lazadaItems;
        private readonly IEnumerable<Item> _shopeeItems;
        private readonly double _torerantRate;

        public ComparedFile(IEnumerable<Item> lazadaItems, IEnumerable<Item> shopeeItems, double torerantRate)
        {
            _lazadaItems = lazadaItems;
            _shopeeItems = shopeeItems;
            _torerantRate = torerantRate;
        }

        public bool Write(string path)
        {
            var shopee = _shopeeItems.ToList();
            var comparedItems = new List<ComparedItem>();
            foreach (var l in _lazadaItems.ToList())
            {
                var acceptedCount = l.Name.Length * _torerantRate;
                var matchedShopee = shopee.FirstOrDefault(s => Compute(s.Name.ToLower(), l.Name.ToLower()) < acceptedCount);
                var comparedItem = new ComparedItem() { LazadaName = l.Name, LazadaPrice = l.Price, LazadaSku = l.SKU };

                if (matchedShopee != null)
                {
                    comparedItem.ShopeeName = matchedShopee.Name;
                    comparedItem.ShopeePrice = matchedShopee.Price;
                }

                comparedItems.Add(comparedItem);
            }

            using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");

                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("Lazada SKU");
                headerRow.CreateCell(1).SetCellValue("Lazada Name");
                headerRow.CreateCell(2).SetCellValue("Shopee Name");
                headerRow.CreateCell(3).SetCellValue("Lazada Price");
                headerRow.CreateCell(4).SetCellValue("Shopee Price");

                for (int i = 0; i < comparedItems.Count; i++)
                {
                    var item = comparedItems[i];
                    var rowtemp = sheet.CreateRow(i + 1);
                    rowtemp.CreateCell(0).SetCellValue(item.LazadaSku);
                    rowtemp.CreateCell(1).SetCellValue(item.LazadaName);
                    rowtemp.CreateCell(2).SetCellValue(item.ShopeeName);
                    rowtemp.CreateCell(3).SetCellValue(item.LazadaPrice.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(4).SetCellValue(item.ShopeePrice.ToString(CultureInfo.InvariantCulture));
                }

                workbook.Write(stream);

                return true;
            }
        }

        public IEnumerable<Item> Read(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sheet1");
            for (int row = 2; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(0).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(4).StringCellValue);

                yield return new Item() { Name = name, Price = price };
            }
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