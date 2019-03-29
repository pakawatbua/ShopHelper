using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;

namespace ShopHelper
{
    internal class UnmatchedNameManager
    {
        private readonly IEnumerable<Item> _sources;
        private readonly IEnumerable<Item> _descs;

        public UnmatchedNameManager(IEnumerable<Item> sources, IEnumerable<Item> descs)
        {
            _sources = sources;
            _descs = descs;
        }

        public void Write(Common.Shop shop, string outputPath)
        {
            switch (shop)
            {
                case Common.Shop.Shopee:
                    WriteShopee(outputPath);
                    break;
                case Common.Shop.Lazada:
                    WriteLazada(outputPath);
                    break;
            }
        }

        private void WriteLazada(string outputPath)
        {
            var results = new List<Item>();

            foreach (var source in _sources)
            {
                var descs = _descs.FirstOrDefault(s => string.Compare(s.Name, source.Name, StringComparison.CurrentCultureIgnoreCase) == 0);
                if (descs == null)
                {
                    var lazName = source.Name;
                    var match90Name = _descs.FirstOrDefault(s => CompareHelper.Compare(s.Name.ToLower(), source.Name.ToLower()) < source.Name.Length * 0.1)?.Name;
                    var match80Name = match90Name == null ? _descs.FirstOrDefault(s => CompareHelper.Compare(s.Name.ToLower(), source.Name.ToLower()) < source.Name.Length * 0.2)?.Name : null;
                    var match70Name = match90Name == null && match80Name == null ? _descs.FirstOrDefault(s => CompareHelper.Compare(s.Name.ToLower(), source.Name.ToLower()) < source.Name.Length * 0.3)?.Name : null;
                    results.Add(new Item() { LazName = lazName, Matched90Name = match90Name, Matched80Name = match80Name, Matched70Name = match70Name });
                }
            }

            using (FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("sheet1");
                var row = 0;

                var headerRow = sheet.CreateRow(row);
                headerRow.CreateCell(0).SetCellValue("lazName");
                headerRow.CreateCell(1).SetCellValue("Matched90Name");
                headerRow.CreateCell(2).SetCellValue("Matched80Name");
                headerRow.CreateCell(3).SetCellValue("Matched70Name");

                foreach (var result in results.OrderByDescending(x => x.Matched90Name).ThenByDescending(x => x.Matched80Name).ThenByDescending(x => x.Matched70Name))
                {
                    var rowtemp = sheet.CreateRow(++row);
                    rowtemp.CreateCell(0).SetCellValue(result.LazName);
                    rowtemp.CreateCell(1).SetCellValue(result.Matched90Name);
                    rowtemp.CreateCell(2).SetCellValue(result.Matched80Name);
                    rowtemp.CreateCell(3).SetCellValue(result.Matched70Name);
                }

                workbook.Write(stream);
            }
        }

        private void WriteShopee(string outputPath)
        {
            throw new NotImplementedException();
        }
    }
}
