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
                // Source longer than desc
                matched =
                descs.Where(s => source.Name.ToLower().Contains(s.Name.ToLower())).ToList();
            }

            if (!matched.Any())
            {
                // Desc longer than source
                matched =
                descs.Where(s => s.Name.ToLower().Contains(source.Name.ToLower())).ToList();
            }

            if (!matched.Any())
            {
                // 90% match
                matched = descs.Where(s => CompareHelper.Compare(s.Name.ToLower(), source.Name.ToLower()) < source.Name.Length * 0.1).ToList();
            }

            if (!matched.Any())
            {
                // 80% match
                matched = descs.Where(s => CompareHelper.Compare(s.Name.ToLower(), source.Name.ToLower()) < source.Name.Length * 0.2).ToList();
            }

            // Not found any
            if (!matched.Any())
            {
                source.Matched = false;
                return source;
            }

            // One one
            if(matched.Count == 1)
            {
                var firstMatch = matched.First();
                if (firstMatch.AltName == null)
                {
                    firstMatch.Matched = true;
                    return firstMatch;
                }

            }

            // For fix missing
            var fixMatched = descs.Where(s => s.AltName != null
            && source.Name.Equals(s.Name, StringComparison.InvariantCultureIgnoreCase)
            && source.SKU.Equals(s.AltName, StringComparison.InvariantCultureIgnoreCase)
            ).FirstOrDefault();

            if(fixMatched != null)
            {
                fixMatched.Matched = true;
                return fixMatched;
            }
            
            // Multiple match
            var altSize = Regex.Match(source.SKU, @"\d+").Value;
            foreach (var alt in matched.Where(s => s.AltName != null))
            {
                alt.AltName = Regex.Match(alt.AltName, @"\d+").Value;
            }

            var altMatched = matched.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(altSize.ToLower(), s.AltName.ToLower())).First();
            altMatched.Matched = true;

            return altMatched;
        }
    }
}
