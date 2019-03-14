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
                        default:
                            return null;
                    }
                default:
                    return null;
            }
        }

        private List<Item> GetLazadaPrice(string path)
        {
            throw new NotImplementedException();
        }

        private List<Item> GeLazadaStock(string path)
        {
            throw new NotImplementedException();
        }

        private List<Item> GetShopeePrice(string path)
        {
            throw new NotImplementedException();
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

                yield return new Item() { Name = name, Price = price , Stock = stock  };
            }
        }
    }
}
