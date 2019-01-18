using System.Collections.Generic;

namespace ShopHelper.Interfaces
{
    interface IFileManager
    {
        IEnumerable<Item> Read(string path);

        bool Write(string path);
    }
}
