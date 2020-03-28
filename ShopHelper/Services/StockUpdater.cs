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
                    Model = tergetStock.SKU,
                    MatchedModel = matched.Matched ? matched.SKU + " : " + matched.Stock : string.Empty,
                    Stock = matched.Matched ? matched.Stock : tergetStock.Stock,
                    Matched = matched.Matched,
                    MultiStocks = matched.MultiStocks,
                    MultiPrices = matched.MultiPrices
                });
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("Name");
                headerRow.CreateCell(1).SetCellValue("Model");
                headerRow.CreateCell(2).SetCellValue("Matched Model");
                headerRow.CreateCell(3).SetCellValue("Stock");
                headerRow.CreateCell(4).SetCellValue("Matched");
                headerRow.CreateCell(5).SetCellValue("MultiStocks");
                headerRow.CreateCell(6).SetCellValue("MultiPrices");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.Name);
                    rowtemp.CreateCell(1).SetCellValue(result.Model);
                    rowtemp.CreateCell(2).SetCellValue(result.MatchedModel);
                    rowtemp.CreateCell(3).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(4).SetCellValue(!result.Matched ? "NO" : string.Empty);
                    rowtemp.CreateCell(5).SetCellValue(result.MultiStocks);
                    rowtemp.CreateCell(6).SetCellValue(result.MultiPrices);
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
                    Model = tergetStock.Model,
                    MatchedModel = matched.Matched ? matched.Model + " : " + matched.Stock : string.Empty,
                    Stock = matched.Matched ? matched.Stock : tergetStock.Stock,
                    Matched = matched.Matched,
                    MultiStocks = matched.MultiStocks,
                    MultiPrices = matched.MultiPrices
                });
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("Name");
                headerRow.CreateCell(1).SetCellValue("Model");
                headerRow.CreateCell(2).SetCellValue("Matched Model");
                headerRow.CreateCell(3).SetCellValue("Stock");
                headerRow.CreateCell(4).SetCellValue("Matched");
                headerRow.CreateCell(5).SetCellValue("MultiStocks");
                headerRow.CreateCell(6).SetCellValue("MultiPrices");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.Name);
                    rowtemp.CreateCell(1).SetCellValue(result.Model);
                    rowtemp.CreateCell(2).SetCellValue(result.MatchedModel);
                    rowtemp.CreateCell(3).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(4).SetCellValue(!result.Matched ? "NO" : string.Empty);
                    rowtemp.CreateCell(5).SetCellValue(result.MultiStocks);
                    rowtemp.CreateCell(6).SetCellValue(result.MultiPrices);
                }

                workbook.Write(stream);
            }
        }
    }
}