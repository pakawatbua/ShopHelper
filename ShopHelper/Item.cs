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
        
        public decimal Cost { get; set; }
        public decimal Sell { get; set; }
        public bool Changed { get; set; }
        public string LazName { get; set; }
        public string ShopeeName { get; set; }
        public string Matched70Name { get; set; }
        public string Matched80Name { get; set; }
        public string Matched90Name { get; set; }
        public string AltName { get; set; }
        public bool Matched { get; set; }
        public decimal LazPrice { get; set; }
        public decimal ShopPrice { get; set; }
        public bool IsOverPrice { get; set; }
        public bool kingTag { get; set; }
        public decimal KingPrice { get; set; }
    }
}