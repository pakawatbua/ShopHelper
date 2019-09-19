using System;
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
        private readonly IEnumerable<Item> _baseStock;
        private readonly IEnumerable<Item> _targetStock;

        public StockUpdater(IEnumerable<Item> baseStock, IEnumerable<Item> targetStock)
        {
            _baseStock = baseStock;
            _targetStock = targetStock;
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

        /// <summary>
        /// If not exist in shopee then lazada stock
        /// If exist shopee no alt name then shopee stock
        /// If exist shopee have alt then pick one
        /// </summary>
        /// <param name="laz"></param>
        /// <param name="shopees"></param>
        /// <returns></returns>
        private int GetMatcherdLazadaNShopee(Item laz, List<Item> shopees)
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
        private Item GetMatchedShopeeNShopee(Item desc, List<Item> sources, ref bool changed)
        {
            var matched = _baseStock.FirstOrDefault(s => string.CompareOrdinal(s.Name, desc.Name) == 0);
            if (matched != null && desc.AltName != null)
            {
                var alt = sources.Where(s => s.Name == matched.Name).OrderBy(s => CompareHelper.Compare(desc.AltName, s.AltName ?? "")).FirstOrDefault();
                if (alt != null)
                {
                    changed = true;
                    return alt;
                }
            }

            if (matched != null && desc.AltName == null)
            {
                changed = true;
                return matched;
            }

            changed = false;
            return desc;
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