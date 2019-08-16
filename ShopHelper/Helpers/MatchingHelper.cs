using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ShopHelper
{
    public static class MatchingHelper
    {
        public enum MatchingType
        {
            LazToSho,
            ShoToSho
        }

        /// <summary>
        /// If not exist in shopee then lazada stock
        /// If exist shopee no alt name then shopee stock
        /// If exist shopee have alt then pick one
        /// </summary>
        /// <param name="laz"></param>
        /// <param name="shopees"></param>
        /// <returns></returns>
        public static Item Match(MatchingType lazToSho, Item source, IEnumerable<Item> descs)
        {
            var matched =
                descs.Where(s => source.Name.Equals(s.Name, StringComparison.InvariantCultureIgnoreCase)).ToList();

            if (!matched.Any())
            {
                source.Matched = false;
                return source;
            }

            var firstMatch = matched.First();
            if (firstMatch.AltName == null)
            {
                firstMatch.Matched = true;
                return firstMatch;
            }

            var altMatched = matched.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(source.SKU.ToLower(), s.AltName.ToLower())).First();
            altMatched.Matched = true;

            if (altMatched.Price > (source.Price * (decimal) 1.5))
            {
                var altSize = Regex.Match(source.SKU, @"\d+").Value;
                foreach (var alt in matched)
                {
                    alt.AltName = Regex.Match(alt.AltName, @"\d+").Value;
                }

                altMatched = matched.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(altSize.ToLower(), s.AltName.ToLower())).First();
                altMatched.Matched = true;
            }

            return altMatched;
        }
    }
}
