using System.Collections.Generic;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper.Interfaces
{
    interface IFileManager
    {
        IEnumerable<Item> Read(string path);

        bool Write(string path);
    }
}
