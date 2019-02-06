using System.Collections.Generic;
using ShopHelper.Interfaces;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace ShopHelper
{
    internal class LazadaBasePriceFile
    {
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

                var name = sheet.GetRow(row).GetCell(0).StringCellValue;
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
