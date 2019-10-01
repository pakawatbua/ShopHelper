using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;
namespace ShopHelper.Services
{
    public class TopSellPriceUpdater
    {
        private readonly IEnumerable<Item> _topSell;
        private readonly IEnumerable<Item> _basePrice;
        private readonly IEnumerable<Item> _cost;

        public TopSellPriceUpdater(IEnumerable<Item> topSell, IEnumerable<Item> basePrice, IEnumerable<Item> cost)
        {
            _topSell = topSell;
            _basePrice = basePrice;
            _cost = cost;
        }

        public void Write(Common.Shop shop, string outputPath)
        {
            switch (shop)
            {
                case Common.Shop.Lazada:
                    WriteLazada(outputPath);
                    break;
            }
        }

        private void WriteLazada(string outputPath)
        {
            var results = new List<Item>();

            foreach (var topsell in _topSell)
            {
                var matched = MatchingHelper.MatchSku(topsell, _basePrice);

                var addedCost = MatchingHelper.Match(matched, _cost);

                results.Add(new Item()
                {
                    SKU = matched.SKU,
                    Price = matched.Price,
                    SalePrice = matched.SalePrice,
                    SaleStartDate = matched.SaleStartDate,
                    SaleEndDate = matched.SaleEndDate,
                    Name = matched.Name,
                    Cost = addedCost.Price,
                    Matched = addedCost.Matched,
                });
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("SKU");
                headerRow.CreateCell(1).SetCellValue("Price");
                headerRow.CreateCell(2).SetCellValue("SalePrice");
                headerRow.CreateCell(3).SetCellValue("SaleStartDate");
                headerRow.CreateCell(4).SetCellValue("SaleEndDate");
                headerRow.CreateCell(5).SetCellValue("Name");
                headerRow.CreateCell(6).SetCellValue("Cost");
                headerRow.CreateCell(7).SetCellValue("Matched");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.SKU);
                    rowtemp.CreateCell(1).SetCellValue(result.Price.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(result.SalePrice.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.SaleStartDate);
                    rowtemp.CreateCell(4).SetCellValue(result.SaleEndDate);
                    rowtemp.CreateCell(5).SetCellValue(result.Name);
                    rowtemp.CreateCell(6).SetCellValue(result.Cost.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(7).SetCellValue(!result.Matched ? "NO" : string.Empty);
                }

                workbook.Write(stream);
            }
        }
    }
}
