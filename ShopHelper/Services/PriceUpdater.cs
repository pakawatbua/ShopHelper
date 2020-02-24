using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper.Services
{
    internal class PriceUpdater
    {
        private readonly List<Item> _baseStock;
        private readonly List<Item> _targetStock;

        public PriceUpdater(IEnumerable<Item> baseStock, IEnumerable<Item> targetStock)
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
                    break;
            }
        }

        private void WriteShopee(string outputPath)
        {
            var results = new List<Item>();
            foreach (var tergetStock in _targetStock)
            {
                var matched = MatchingHelper.Match(tergetStock, _baseStock);

                var result = new Item()
                {
                    Name = tergetStock.Name,
                    PriceBYM = tergetStock.Price,
                    PricePlank = matched.Matched ? matched.Price : 0,
                    Price = matched.Matched ? matched.Price : tergetStock.Price,
                    Matched = matched.Matched
                };

                results.Add(result);

                if (result.Matched)
                {
                    if (result.PricePlank - result.PriceBYM >= 100)
                    {
                        result.Description += "Plank over BYM. Wrong Price!!";
                    }
                    else if (result.PriceBYM - result.PricePlank >= 100)
                    {
                        result.Description += "BYM over Plank. Wrong Price!!";
                    }
                }  
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("Name");
                headerRow.CreateCell(1).SetCellValue("PriceBYM");
                headerRow.CreateCell(2).SetCellValue("PricePlank");
                headerRow.CreateCell(3).SetCellValue("Price");
                headerRow.CreateCell(4).SetCellValue("Matched");
                headerRow.CreateCell(5).SetCellValue("Description");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.Name);
                    rowtemp.CreateCell(1).SetCellValue(result.PriceBYM.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(result.PricePlank.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.Price.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(4).SetCellValue(result.Matched);
                    rowtemp.CreateCell(5).SetCellValue(result.Description);
                }

                workbook.Write(stream);
            }
        }
    }
}
