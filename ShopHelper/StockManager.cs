using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;

namespace ShopHelper
{
    internal class StockManager
    {
        private readonly IEnumerable<Item> _sources;
        private readonly IEnumerable<Item> _descs;

        public StockManager(IEnumerable<Item> sources, IEnumerable<Item> descs)
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

        private void WriteLazada(string outputPath)
        {
            var results = new List<Item>();

            foreach (var laz in _sources)
            {
                var shopee = _descs.Where(
                    s => string.Compare(s.Name, laz.Name, StringComparison.InvariantCultureIgnoreCase) == 0).ToList();

                var builder = new StringBuilder();

                if (shopee.Count > 1)
                {
                    shopee.ForEach(x => builder.Append($"{x.AltName} : {x.Stock}, "));
                }

                results.Add(new Item()
                {
                    LazName = laz.Name,
                    SKU = laz.SKU,
                    Stock = GetMatcherdStock(laz, shopee),
                    Changed = shopee.Count != 0,
                    AltName = builder.ToString()
                });
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("lazName");
                headerRow.CreateCell(1).SetCellValue("SKU");
                headerRow.CreateCell(2).SetCellValue("Stock");
                headerRow.CreateCell(3).SetCellValue("Changed");
                headerRow.CreateCell(4).SetCellValue("AltName");

                foreach (var result in results)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.LazName);
                    rowtemp.CreateCell(1).SetCellValue(result.SKU);
                    rowtemp.CreateCell(2).SetCellValue(result.Stock.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(3).SetCellValue(result.Changed);
                    rowtemp.CreateCell(4).SetCellValue(result.AltName);
                }

                workbook.Write(stream);
            }
        }

        /// <summary>
        /// If not exist in shopee then lazada stock
        /// If exist shopee no alt name then shopee stock
        /// If exist shopee have alt then pick one
        /// </summary>
        /// <param name="laz"></param>
        /// <param name="shopees"></param>
        /// <returns></returns>
        private int GetMatcherdStock(Item laz, List<Item> shopees)
        {
            if (!shopees.Any()) return laz.Stock;

            if (shopees.First().AltName == null) return shopees.First().Stock;

            return shopees.OrderBy(s => CompareHelper.Compare(laz.SKU, s.AltName)).First().Stock;
        }

        private void WriteShopee(string outputPath)
        {
            var results = new List<Item>();
            foreach (var dStock in _descs)
            {
                var compared = _sources.FirstOrDefault(s => string.CompareOrdinal(s.Name, dStock.Name) == 0);
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