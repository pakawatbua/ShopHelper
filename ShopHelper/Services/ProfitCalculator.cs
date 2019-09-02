using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper
{
    internal class ProfitCalculator
    {
        private readonly IEnumerable<Item> _cost;
        private readonly IEnumerable<Item> _sell;

        public ProfitCalculator(IEnumerable<Item> cost, IEnumerable<Item> sell)
        {
            _cost = cost;
            _sell = sell;
        }

        public void Write(Common.Shop shop, string outputPath)
        {
            var results = new List<Item>();

            foreach (var sell in _sell)
            {
                try
                {
                    var matched = MatchingHelper.Match(sell, _cost);

                var sku = shop == Common.Shop.Lazada ?
                    sell.SKU : 
                    sell.AltName ;

                
                    results.Add(new Item()
                    {
                        LazName = sell.Name,
                        SKU = sku,
                        Sell = sell.Price,
                        Cost = matched.Matched ? matched.Price : 0,
                        Matched = matched.Matched,
                        IsOverPrice = matched.Matched ? matched.Price > sell.Price : false,
                        Amount = sell.Amount,
                        CostType = matched.Matched ? matched.CostType : string.Empty
                    });
                }
                catch (Exception)
                {
                    throw;
                }
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("profit");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("lazName");
                headerRow.CreateCell(1).SetCellValue("SKU");
                headerRow.CreateCell(2).SetCellValue("Sell");
                headerRow.CreateCell(3).SetCellValue("Cost");
                headerRow.CreateCell(4).SetCellValue("Not Matched");
                headerRow.CreateCell(5).SetCellValue("Over Price");
                headerRow.CreateCell(6).SetCellValue("Type");
                headerRow.CreateCell(7).SetCellValue("Amount");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.LazName);
                    rowtemp.CreateCell(1).SetCellValue(result.SKU);
                    rowtemp.CreateCell(2).SetCellValue(result.Sell.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.Cost.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(4).SetCellValue(result.Matched ? "" : "NO");
                    rowtemp.CreateCell(5).SetCellValue(result.IsOverPrice ? "YES" : "");
                    rowtemp.CreateCell(6).SetCellValue(result.CostType);
                    rowtemp.CreateCell(7).SetCellValue(result.Amount.ToString(CultureInfo.InvariantCulture));
                }

                workbook.Write(stream);
            }
        }
    }
}
