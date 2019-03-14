namespace ShopHelper
{
    public class Item
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public decimal BasePrice { get; set; }
        public string SKU { get; set; }
        public decimal Profit { get; set; }
        public bool Changed { get; set; }
    }
}