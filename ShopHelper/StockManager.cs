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

            try
            {
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
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
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

            return shopees.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(laz.SKU, s.AltName)).First().Stock;
        }

        /// <summary>
        /// IF alt match then matched
        /// IF name match then matched
        /// IF not then desc
        /// </summary>
        /// <param name="desc"></param>
        /// <param name="sources"></param>
        /// <param name="changed"></param>
        private Item GetMatcherdItem(Item desc, List<Item> sources, ref bool changed)
        {
            if (desc.AltName != null)
            {
                changed = true;
                var alt = sources.OrderBy(s => CompareHelper.Compare(desc.AltName, s.AltName?? "")).FirstOrDefault();
                if (alt != null)
                {
                    return alt;
                }
            }

            var matched = _sources.FirstOrDefault(s => string.CompareOrdinal(s.Name, desc.Name) == 0);

            if (matched != null)
            {
                changed = true;
                return matched;
            }

            changed = false;
            return desc;
        }

        private void WriteShopee(string outputPath)
        {
            try
            {

            
            var results = new List<Item>();
            foreach (var dStock in _descs)
            {
                bool changed = true;
                var matched = GetMatcherdItem(dStock, _sources.ToList(), ref changed);
                var name = dStock.Name;
                var stock = matched.Stock;
                var price = matched.Price;

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
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }
    }
}