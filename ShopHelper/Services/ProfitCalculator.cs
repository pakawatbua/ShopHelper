using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;

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
            switch (shop)
            {
                case Common.Shop.Shopee:
                    WriteShopee(outputPath);
                    break;
                case Common.Shop.Lazada:
                    WriteLazada(outputPath);
                    break;
            }
        }

        private void WriteLazada(string outputPath)
        {
            var results = new List<Item>();

            foreach (var sell in _sell)
            {
                var matched = MatchingHelper.Match(MatchingHelper.MatchingType.LazToSho, sell, _cost);

                var builder = new StringBuilder();

                results.Add(new Item()
                {
                    LazName = sell.Name,
                    SKU = sell.SKU,
                    Sell = sell.Price,
                    Cost = matched.Matched ? matched.Price : sell.Price,
                });
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

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.LazName);
                    rowtemp.CreateCell(1).SetCellValue(result.SKU);
                    rowtemp.CreateCell(2).SetCellValue(result.Sell.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.Cost.ToString(CultureInfo.InvariantCulture));
                }

                workbook.Write(stream);
            }
        }

        private void WriteShopee(string outputPath)
        {
            throw new NotImplementedException();
        }
    }
}
