using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ShopHelper.Models;
using NPOI.XSSF.UserModel;

namespace ShopHelper
{
    public class ProfitFile
    {
        private readonly IEnumerable<Item> _lazadaSellList;
        private readonly IEnumerable<Item> _lazadaBasePrice;

        public ProfitFile(IEnumerable<Item> lazadaSellList, IEnumerable<Item> lazadaBasePrice)
        {
            _lazadaSellList = lazadaSellList;
            _lazadaBasePrice = lazadaBasePrice;
        }

        public bool Write(string path)
        {
            foreach (var selled in _lazadaSellList)
            {
                selled.BasePrice = _lazadaBasePrice.First(x => x.Name == selled.Name).Price;
                selled.Profit = selled.Price - selled.BasePrice;
            }

            using(FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("Name");
                headerRow.CreateCell(1).SetCellValue("Price");
                headerRow.CreateCell(2).SetCellValue("Profit");


                foreach (var selled in _lazadaSellList)
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(selled.Name);
                    rowtemp.CreateCell(1).SetCellValue(selled.Price.ToString(CultureInfo.InvariantCulture));
                    rowtemp.CreateCell(2).SetCellValue(selled.Profit.ToString(CultureInfo.InvariantCulture));
                }

                workbook.Write(stream);
            }

            return true;
        }
    }
}
