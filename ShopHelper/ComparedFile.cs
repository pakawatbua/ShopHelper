using System.Collections.Generic;
using ShopHelper.Interfaces;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;

namespace ShopHelper
{
    public class ComparedFile
    {
        List<ComparedItem> _comparedItems;

        public ComparedFile()
        {

        }

        public ComparedFile(IEnumerable<ComparedItem> comparedItems)
        {
            _comparedItems = comparedItems.ToList();
        }

        public bool Write(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");

                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("Lazada SKU");
                headerRow.CreateCell(1).SetCellValue("Lazada Name");
                headerRow.CreateCell(2).SetCellValue("Shopee Name");
                headerRow.CreateCell(3).SetCellValue("Lazada Price");
                headerRow.CreateCell(4).SetCellValue("Shopee Price");

                for (int i = 0; i < _comparedItems.Count; i++)
                {
                    var item = _comparedItems[i];
                    var rowtemp = sheet.CreateRow(i + 1);
                    rowtemp.CreateCell(0).SetCellValue(item.LazadaSku);
                    rowtemp.CreateCell(1).SetCellValue(item.LazadaName);
                    rowtemp.CreateCell(2).SetCellValue(item.ShopeeName);
                    rowtemp.CreateCell(3).SetCellValue(item.LazadaPrice.ToString());
                    rowtemp.CreateCell(4).SetCellValue(item.ShopeePrice.ToString());
                }

                workbook.Write(stream);

                return true;
            }
        }

        public IEnumerable<Item> Read(string path)
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

                var name = sheet.GetRow(row).GetCell(0).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(4).StringCellValue);

                yield return new Item() { Name = name, Price = price };
            }
        }
    }
}