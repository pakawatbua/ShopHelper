using System.Collections.Generic;
using ShopHelper.Interfaces;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;

namespace ShopHelper
{
    internal class ShopeeFile : IFileManager
    {
        public ShopeeFile()
        {

        }

        public IEnumerable<Item> Read(string path)
        {
            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sheet1");
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) == null) continue;

                var name = sheet.GetRow(row).GetCell(2).StringCellValue;
                var price = decimal.Parse(sheet.GetRow(row).GetCell(6).StringCellValue);

                yield return new Item() { Name = name, Price = price };
            }
        }

        public bool Write(string path)
        {
            throw new System.NotImplementedException();
        }
    }
}