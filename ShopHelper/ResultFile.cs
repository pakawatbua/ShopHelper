using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ShopHelper
{
    internal class ResultFile
    {
        List<ComparedItem> _comparedItems;
        public ResultFile(IEnumerable<ComparedItem> comparedItems)
        {
            _comparedItems = comparedItems.ToList();
        }

        internal bool Write(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");

                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("Lazada Name");
                headerRow.CreateCell(1).SetCellValue("Shopee Name");
                headerRow.CreateCell(2).SetCellValue("Lazada Price");
                headerRow.CreateCell(3).SetCellValue("Shopee Price");

                for (int i = 0; i < _comparedItems.Count; i++)
                {
                    var item = _comparedItems[i];
                    var rowtemp = sheet.CreateRow(i + 1);
                    rowtemp.CreateCell(0).SetCellValue(item.LazadaName);
                    rowtemp.CreateCell(1).SetCellValue(item.ShopeeName);
                    rowtemp.CreateCell(2).SetCellValue(item.LazadaPrice.ToString());
                    rowtemp.CreateCell(3).SetCellValue(item.ShopeePrice.ToString());
                }

                workbook.Write(stream);

                return true;
            }
        }
    }
}