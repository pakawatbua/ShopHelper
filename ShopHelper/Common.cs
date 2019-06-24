using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShopHelper
{
    public static class Common
    {
        public enum Shop
        {
            Lazada,
            Shopee
        }

        public enum Type
        {
            Stock,
            Price,
            Product,
            CampaignPrice
        }
    }
}
