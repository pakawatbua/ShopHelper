using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper.Services
{
    internal class Merger
    {
        private readonly List<Item> _prices;
        private readonly List<Item> _stocks;

        public Merger(IEnumerable<Item> prices, IEnumerable<Item> stocks)
        {
            _prices = prices.ToList();
            _stocks = stocks.ToList();
        }

        public void Write(Common.Shop shop, string outputPath)
        {
            switch (shop)
            {
                case Common.Shop.BYM:
                    break;
                case Common.Shop.Lazada:
                    WriteLazada(outputPath);
                    break;
            }
        }

        private void WriteLazada(string outputPath)
        {
            var results = new List<Item>();

            foreach (var stock in _stocks)
            {
                var price = _prices.First(p => p.SKU == stock.SKU);

                results.Add(new Item()
                {
                   SKU = stock.SKU,
                   Stock = stock.Stock,
                   Name = stock.Name,
                   Price = price.Price
                });
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("SKU");
                headerRow.CreateCell(1).SetCellValue("Stock");
                headerRow.CreateCell(2).SetCellValue("Name");
                headerRow.CreateCell(3).SetCellValue("Price");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.SKU);
                    rowtemp.CreateCell(1).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(result.Name);
                    rowtemp.CreateCell(3).SetCellValue(result.Price.ToString(CultureInfo.InvariantCulture));
                }

                workbook.Write(stream);
            }
        }

    }
}
