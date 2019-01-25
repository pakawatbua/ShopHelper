using System.Collections.Generic;
using ShopHelper.Interfaces;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ShopHelper
{
    internal class LazadaFile : IFileManager
    {
        public LazadaFile()
        {
        }

        public IEnumerable<Item> Read(string path)
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("template");
            for (int row = 2; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(5).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(1).StringCellValue);

                yield return new Item() { Name = name, Price = price };
            }
        }

        public bool Write(string path)
        {
            throw new System.NotImplementedException();
        }
    }
}