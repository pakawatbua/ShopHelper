using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;
namespace ShopHelper
{
    internal class PriceManager
    {
        private readonly IEnumerable<Item> _sources;
        private readonly IEnumerable<Item> _descs;

        public PriceManager(IEnumerable<Item> sources, IEnumerable<Item> descs)
        {
            _sources = sources;
            _descs = descs;
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

        private void WriteShopee(string outputPath)
        {
            var results = new List<Item>();
            foreach (var dStock in _descs)
            {
                var compared = _sources.FirstOrDefault(s => string.CompareOrdinal(s.Name, dStock.Name) == 0);
                var name = dStock.Name;
                var stock = compared?.Stock ?? dStock.Stock;
                var price = compared?.Price -1 ?? dStock.Price;
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

        private void WriteLazada(string outputPath)
        {
            var shopee = _sources;
            var lazada = _descs;

            var results = new List<Item>();

            foreach (var laz in lazada)
            {
                var matched = MatchingHelper.Match(MatchingHelper.MatchingType.LazToSho, laz, shopee);

                results.Add(new Item()
                {
                    LazName = laz.Name,
                    SKU = laz.SKU,
                    LazPrice = laz.Price,
                    ShopPrice = matched.Matched ? matched.Price : 0
                    //IsOverPrice = (matched.Matched ? matched.Price : 0) > laz.Price ? "Yes" : "",
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
                headerRow.CreateCell(2).SetCellValue("LazPrice");
                headerRow.CreateCell(3).SetCellValue("ShopPrice");
                headerRow.CreateCell(4).SetCellValue("IsOverPrice");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.LazName);
                    rowtemp.CreateCell(1).SetCellValue(result.SKU);
                    rowtemp.CreateCell(2).SetCellValue(result.LazPrice.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.ShopPrice.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(4).SetCellValue(result.IsOverPrice);
                }

                workbook.Write(stream);
            }
        }
    }
}
