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
                case Common.Shop.BYM:
                case Common.Shop.Shopee:
                    switch (_type)
                    {
                        case Common.Type.Stock:
                            return GetShopeeStock(path);
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
                            return GetLazadaStock(path);
                        case Common.Type.UpgradStock:
                            return GetLazadaUpgradeStock(path);
                        case Common.Type.Price:
                            return GetLazadaPrice(path);
                        case Common.Type.Product:
                            return GetLazadaProduct(path);
                        case Common.Type.Sell:
                            return GetLazadaSell(path);
                        case Common.Type.TopSellPrice:
                            return GetLazadaTop100Items(path);
                        case Common.Type.PriceTemplate:
                            return GetLazadaPriceTemplate(path);
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
                if (sheet.GetRow(row) == null) break;

                string name;
                string sku;
                decimal price;

                name = sheet.GetRow(row).GetCell(4).StringCellValue;
                sku = sheet.GetRow(row).GetCell(5).StringCellValue;
                price = decimal.Parse(sheet.GetRow(row).GetCell(7).NumericCellValue.ToString(CultureInfo.InvariantCulture));

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

        private IEnumerable<Item> GetLazadaStock(string path)
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
                var stock = int.Parse(sheet.GetRow(row).GetCell(1).ToString());
                var name = sheet.GetRow(row).GetCell(2).StringCellValue;

                yield return new Item() { Name = name, Stock = stock, SKU = sku };
            }
        }

        private IEnumerable<Item> GetLazadaUpgradeStock(string path)
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
                var stock = int.Parse(sheet.GetRow(row).GetCell(1).ToString());
                var price = decimal.Parse(sheet.GetRow(row).GetCell(3).ToString());

                yield return new Item() { Name = name, Stock = stock, SKU = sku, Price = price };
            }
        }

        private IEnumerable<Item> GetLazadaTop100Items(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("สินค้า");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(0).StringCellValue;
                var sku = sheet.GetRow(row).GetCell(1).StringCellValue;

                yield return new Item() { Name = name, SKU = sku };
            }
        }

        private IEnumerable<Item> GetLazadaPriceTemplate(string path)
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
                var salePrice = decimal.Parse(sheet.GetRow(row).GetCell(2).StringCellValue);
                var saleStartDate = sheet.GetRow(row).GetCell(3).StringCellValue;
                var saleEndDate = sheet.GetRow(row).GetCell(4).StringCellValue;
                var name = sheet.GetRow(row).GetCell(5).StringCellValue;

                yield return new Item()
                {
                    SKU = sku,
                    Price = price,
                    SalePrice = salePrice,
                    SaleStartDate = saleStartDate,
                    SaleEndDate = saleEndDate,
                    Name = name
                };
            }
        }

        #endregion

        #region Shopee

        private IEnumerable<Item> GetShopeePrice(string path)
        {
            return GetShopee(path);
        }

        private IEnumerable<Item> GetShopeeStock(string path)
        {
            return GetShopee(path);
        }

        private IEnumerable<Item> GetShopee(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sheet1");
            for (int row = 3; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                string name;
                decimal price;
                int stock;
                string altName;

                try
                {
                    name = sheet.GetRow(row).GetCell(1).StringCellValue;
                    price = decimal.Parse(sheet.GetRow(row).GetCell(6).StringCellValue);
                    stock = int.Parse(sheet.GetRow(row).GetCell(7).StringCellValue);
                    altName = sheet.GetRow(row).GetCell(3)?.StringCellValue;
                }
                catch (System.Exception ex)
                {

                    throw;
                }

                yield return new Item() { Name = name, Price = price, Stock = stock, Model = altName };
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

                string name;
                decimal price;
                string altName;
                double? giftPrice;
                double? buyPrice;
                double? king20Price;
                double? king10Price;
                string costType;

                try
                {
                    name = sheet.GetRow(row).GetCell(2).StringCellValue;
                    price = decimal.Parse(sheet.GetRow(row).GetCell(15).NumericCellValue.ToString(CultureInfo.InvariantCulture));
                    altName = sheet.GetRow(row).GetCell(5)?.StringCellValue;
                    giftPrice = sheet.GetRow(row).GetCell(11)?.NumericCellValue;
                    buyPrice = sheet.GetRow(row).GetCell(12)?.NumericCellValue;
                    king20Price = sheet.GetRow(row).GetCell(13)?.NumericCellValue;
                    king10Price = sheet.GetRow(row).GetCell(14)?.NumericCellValue;
                    costType = (buyPrice != 0 && buyPrice != null) ? "Buy" :
                        (giftPrice != 0 && giftPrice != null) ? "Gift" :
                        (king10Price != 0 && king10Price != null) ? "King10" :
                        (king20Price != 0 && king20Price != null) ? "King20" :
                        string.Empty;
                }
                catch (System.Exception)
                {

                    throw;
                }

                yield return new Item() { Name = name, Price = price, Model = altName, CostType = costType };
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

                yield return new Item() { Name = name, Price = price, Model = altName, Amount = amount };
            }
        }

        #endregion
    }
}
