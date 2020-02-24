using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper
{
    internal class StockUpdater
    {
        private readonly List<Item> _baseStock;
        private readonly List<Item> _targetStock;

        public StockUpdater(IEnumerable<Item> baseStock, IEnumerable<Item> targetStock)
        {
            _baseStock = baseStock.ToList();
            _targetStock = targetStock.ToList();
        }

        public void Write(Common.Shop shop, string outputPath)
        {
            switch (shop)
            {
                case Common.Shop.BYM:
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

            foreach (var tergetStock in _targetStock)
            {
                var matched = MatchingHelper.Match(tergetStock, _baseStock);

                results.Add(new Item()
                {
                    Name = tergetStock.Name,
                    Stock = matched.Matched ? matched.Stock : tergetStock.Stock,
                    Price = matched.Matched ? matched.Price : tergetStock.Price,
                    Matched = matched.Matched
                });
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
                headerRow.CreateCell(3).SetCellValue("Matched");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.Name);
                    rowtemp.CreateCell(1).SetCellValue(result.Price.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(!result.Matched ? "NO" : string.Empty);
                }

                workbook.Write(stream);
            }
        }

        private void WriteShopee(string outputPath)
        {
            var results = new List<Item>();
            foreach (var tergetStock in _targetStock)
            {
                var matched = MatchingHelper.Match(tergetStock, _baseStock);

                results.Add(new Item()
                {
                    Name = tergetStock.Name,
                    Stock = matched.Matched ? matched.Stock : tergetStock.Stock,
                    Price = matched.Matched ? matched.Price : tergetStock.Price,
                    Matched = matched.Matched
                });
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
                headerRow.CreateCell(3).SetCellValue("Matched");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.Name);
                    rowtemp.CreateCell(1).SetCellValue(result.Price.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.Matched);
                }

                workbook.Write(stream);
            }
        }
    }
}