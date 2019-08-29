﻿using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper
{
    public class File
    {
        private readonly Common.Shop _shop;
        private readonly Common.Type _type;

        public File(Common.Shop shop, Common.Type type)
        {
            _shop = shop;
            _type = type;
        }

        public IEnumerable<Item> Read(string path)
        {
            switch (_shop)
            {
                case Common.Shop.Shopee:
                    switch (_type)
                    {
                        case Common.Type.Stock:
                        case Common.Type.Price:
                            return GetShopeePrice(path);
                        case Common.Type.Cost:
                            return GetShopeeCost(path);
                        case Common.Type.Sell:
                            return GetShopeeSell(path);
                        default:
                            return null;
                    }
                case Common.Shop.Lazada:
                    switch (_type)
                    {
                        case Common.Type.Stock:
                            return GeLazadaStock(path);
                        case Common.Type.Price:
                            return GetLazadaPrice(path);
                        case Common.Type.Product:
                            return GetLazadaProduct(path);
                        case Common.Type.Sell:
                            return GetLazadaSell(path);
                        default:
                            return null;
                    }
                default:
                    return null;
            }
        }

        #region Lazada

        private IEnumerable<Item> GetLazadaSell(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sell");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(4).StringCellValue;
                var sku = sheet.GetRow(row).GetCell(5).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(7).NumericCellValue.ToString(CultureInfo.InvariantCulture));

                yield return new Item() { Name = name, Price = price, SKU = sku };
            }
        }

        private IEnumerable<Item> GetLazadaProduct(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("Sheet1");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var sku = sheet.GetRow(row).GetCell(4).StringCellValue;

                yield return new Item() { SKU = sku };
            }
        }

        private IEnumerable<Item> GetLazadaPrice(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("template");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var sku = sheet.GetRow(row).GetCell(0).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(1).StringCellValue);
                var name = sheet.GetRow(row).GetCell(5).StringCellValue;

                yield return new Item() { Name = name, Price = price, SKU = sku };
            }
        }

        private IEnumerable<Item> GeLazadaStock(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("template");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(2).StringCellValue;
                var sku = sheet.GetRow(row).GetCell(0).StringCellValue;
                var stock = int.Parse(sheet.GetRow(row).GetCell(1).StringCellValue);

                yield return new Item() { Name = name, Stock = stock, SKU = sku };
            }
        }

        #endregion

        #region Shopee

        private IEnumerable<Item> GetShopeePrice(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sheet1");
            for (int row = 2; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(2).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(6).StringCellValue);
                var stock = int.Parse(sheet.GetRow(row).GetCell(7).StringCellValue);
                var altName = sheet.GetRow(row).GetCell(5)?.StringCellValue;

                yield return new Item() { Name = name, Price = price, Stock = stock, AltName = altName };
            }
        }

        private IEnumerable<Item> GetShopeeCost(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sheet1");
            for (int row = 2; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(2).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(15).NumericCellValue.ToString(CultureInfo.InvariantCulture));
                var altName = sheet.GetRow(row).GetCell(5)?.StringCellValue;

                var kingPriceText = sheet.GetRow(row).GetCell(10);
                var kingPrice = kingPriceText != null ? decimal.Parse(sheet.GetRow(row).GetCell(10).NumericCellValue.ToString(CultureInfo.InvariantCulture)) : 0;
                var kingTag = kingPrice != 0;

                yield return new Item() { Name = name, Price = price, AltName = altName, KingPrice = kingPrice, kingTag = kingTag };
            }
        }

        private IEnumerable<Item> GetShopeeSell(string path)
        {

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("orders");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {

                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(13).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(17).StringCellValue);
                var altName = sheet.GetRow(row).GetCell(15)?.StringCellValue;
                var amount = int.Parse(sheet.GetRow(row).GetCell(18).NumericCellValue.ToString(CultureInfo.InvariantCulture));

                yield return new Item() { Name = name ,Price = price, AltName = altName, Amount = amount };
            }
        }

        #endregion
    }
}
