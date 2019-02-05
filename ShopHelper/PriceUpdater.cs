using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ShopHelper
{
    internal class PriceUpdater
    {
        public PriceUpdater()
        {
        }

        //public bool Update(IEnumerable<Item> comparedItems, string path)
        //{
        //    XSSFWorkbook hssfwb;
        //    using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
        //    {
        //        hssfwb = new XSSFWorkbook(file);
        //    }

        //    ISheet sheet = hssfwb.GetSheet("sheet1");
        //    for (int row = 2; row <= sheet.LastRowNum; row++)
        //    {
        //        if (sheet.GetRow(row) == null) continue;

        //        var name = sheet.GetRow(row).GetCell(2).StringCellValue;
        //        var price = decimal.Parse(sheet.GetRow(row).GetCell(6).StringCellValue);

        //        yield return new Item() { Name = name, Price = price };
        //    }

        //    return true;
        //}
    }
}