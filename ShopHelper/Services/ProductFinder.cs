using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NPOI.XSSF.UserModel;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper
{
    internal class ProductFinder
    {
        private readonly IEnumerable<Item> _list;
        private readonly IEnumerable<Item> _base;

        public ProductFinder(IEnumerable<Item> list, IEnumerable<Item> @base)
        {
            _list = list;
            _base = @base;
        }
    }
}
