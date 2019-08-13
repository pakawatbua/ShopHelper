using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
                            return GetShopeeStock(path);
                        case Common.Type.Price:
                            return GetShopeePrice(path);
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
                        default:
                            return null;
                    }
                default:
                    return null;
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

        private List<Item> GetLazadaPrice(string path)
        {
            throw new NotImplementedException();
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

                yield return new Item() { Name = name, Stock = stock, SKU = sku};
            }
        }

        private IEnumerable<Item> GetShopeePrice(string path)
        {
            return GetShopeeStock(path);
        }

        private IEnumerable<Item> GetShopeeStock(string path)
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
    }
}
