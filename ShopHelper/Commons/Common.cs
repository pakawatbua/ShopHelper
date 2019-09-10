using System.ComponentModel;
using System.Runtime.Serialization;

namespace ShopHelper.Commons
{
    public static class Common
    {
        public enum Shop
        {
            Lazada,
            Shopee,
            BYM
        }

        public enum Type
        {
            Stock,
            Price,
            Product,
            CampaignPrice,
            Sell,
            Cost
        }

        public enum Function
        {
            ProfitcalLaz,
            ProfitcalSho,
            UpdateStockBYM,
            UpdateStockLaz
        }
    }
}
