using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;

namespace ShopHelper
{
    internal class StockManager
    {
        private readonly IEnumerable<Item> _sourceStocks;
        private readonly IEnumerable<Item> _descStocks;

        public StockManager(IEnumerable<Item> sourceStocks, IEnumerable<Item> descStocks)
        {
            _sourceStocks = sourceStocks;
            _descStocks = descStocks;
        }

        public void Write(Common.Shop shop, string outputPath)
        {

            var test1 = _sourceStocks.ToList();
            var test2 = _descStocks.ToList();

            var results = new List<Item>();
            foreach (var dStock in _descStocks)
            {
                var compared = _sourceStocks.FirstOrDefault(s => string.CompareOrdinal(s.Name, dStock.Name) == 0);
                var name = dStock.Name;
                var stock = compared?.Stock ?? dStock.Stock;
                var price = compared?.Price ?? dStock.Price;
                var changed = compared != null;

                results.Add(new Item() { Name = name, Stock = stock, Price = price, Changed = changed });
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("Name");
                headerRow.CreateCell(1).SetCellValue("Price");
                headerRow.CreateCell(2).SetCellValue("Stock");
                headerRow.CreateCell(3).SetCellValue("Changed");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.Name);
                    rowtemp.CreateCell(1).SetCellValue(result.Price.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.Changed);
                }

                workbook.Write(stream);
            }
        }
    }
}