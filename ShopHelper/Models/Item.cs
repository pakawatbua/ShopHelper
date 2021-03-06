﻿using static ShopHelper.Commons.Common;

namespace ShopHelper.Models
{
    public class Item
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public decimal PriceBYM { get; set; }
        public decimal PricePlank { get; set; }
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
        public string Model { get; set; }
        public string MatchedModel { get; set; }
        public bool Matched { get; set; }
        public decimal LazPrice { get; set; }
        public decimal ShopPrice { get; set; }
        public bool IsOverPrice { get; set; }
        public int Amount { get; set; }
        public string CostType { get; set; }
        public decimal SalePrice { get; set; }
        public string SaleStartDate { get; set; }
        public string SaleEndDate { get; set; }
        public string Description { get; set; }
        public string MultiStocks { get; set; }
        public string MultiPrices { get; set; }
    }
}