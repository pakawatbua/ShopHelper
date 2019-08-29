using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper
{
    internal class CampaignManager
    {
        private readonly IEnumerable<Item> _top100Products;

        public CampaignManager(IEnumerable<Item> top100Products)
        {
            _top100Products = top100Products;
        }

        public void Write(Common.Shop shop, string outputPath, string sourcePath)
        {
            switch (shop)
            {
                case Common.Shop.Shopee:
                    WriteShopee(outputPath);
                    break;
                case Common.Shop.Lazada:
                    WriteLazada(outputPath, sourcePath);
                    break;
            }
        }

        private void WriteLazada(string outputPath, string sourcePath)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
                ISheet sheetRead = hssfwb.GetSheet("null1");

                using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
                {
                    var workbook = new XSSFWorkbook();
                    var sheet = workbook.CreateSheet("null1");
                    var writeRow = 0;

                    for (int row = 1; row <= sheetRead.LastRowNum; row++)
                    {
                        if (sheetRead.GetRow(row) == null) continue;

                        var sku = sheetRead.GetRow(row).GetCell(0).StringCellValue;
                        var top100Sku = _top100Products.Select(s => s.SKU).ToList();
                        if (top100Sku.Contains(sku))
                        {
                            var rowtemp = sheet.CreateRow(++writeRow);
                            rowtemp.CreateCell(0).SetCellValue(sheetRead.GetRow(row).GetCell(0).StringCellValue);
                            rowtemp.CreateCell(1).SetCellValue(sheetRead.GetRow(row).GetCell(1).StringCellValue);
                            rowtemp.CreateCell(2).SetCellValue(sheetRead.GetRow(row).GetCell(2).StringCellValue);
                            rowtemp.CreateCell(3).SetCellValue(sheetRead.GetRow(row).GetCell(3).StringCellValue);
                            rowtemp.CreateCell(4).SetCellValue(sheetRead.GetRow(row).GetCell(4).StringCellValue);
                            rowtemp.CreateCell(5).SetCellValue(sheetRead.GetRow(row).GetCell(5).StringCellValue);
                            rowtemp.CreateCell(6).SetCellValue(sheetRead.GetRow(row).GetCell(6).StringCellValue);
                            rowtemp.CreateCell(7).SetCellValue(sheetRead.GetRow(row).GetCell(7).StringCellValue);
                            rowtemp.CreateCell(8).SetCellValue(sheetRead.GetRow(row).GetCell(8).StringCellValue);
                            rowtemp.CreateCell(9).SetCellValue(sheetRead.GetRow(row).GetCell(9).StringCellValue);
                            rowtemp.CreateCell(10).SetCellValue(sheetRead.GetRow(row).GetCell(10).StringCellValue);
                            rowtemp.CreateCell(11).SetCellValue(sheetRead.GetRow(row).GetCell(11).StringCellValue);
                        }
                    }

                    workbook.Write(stream);
                }
            }
        }

        private void WriteShopee(string outputPath)
        {
            throw new NotImplementedException();
        }
    }
}
