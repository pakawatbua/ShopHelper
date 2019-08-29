using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ShopHelper.Commons;
using ShopHelper.Models;

namespace ShopHelper
{
    public static class MatchingHelper
    {
        public static Item Match( Item target, IEnumerable<Item> set)
        {
            var matched =
                set.Where(s => target.Name.ToLower() == s.Name.ToLower()).ToList();

            if (!matched.Any())
                // target longer than set 
                matched =
                set.Where(s => target.Name.ToLower().Contains(s.Name.ToLower())).ToList();

            if (!matched.Any())
                // set longer than targert
                matched =
                set.Where(s => s.Name.ToLower().Contains(target.Name.ToLower())).ToList();

            if (!matched.Any())
            {
                // 90% matched
                matched = set.Where(s => CompareHelper.Compare(s.Name.ToLower(), target.Name.ToLower()) < target.Name.Length * 0.1).ToList();
            }

            if (!matched.Any())
            {
                // 80% matched
                matched = set.Where(s => CompareHelper.Compare(s.Name.ToLower(), target.Name.ToLower()) < target.Name.Length * 0.2).ToList();
            }

            // Not found
            if (!matched.Any())
            {
                target.Matched = false;
                return target;
            }

            // One only no altname
            if (matched.Count == 1)
            {
                var firstMatch = matched.First();
                if (string.IsNullOrEmpty(firstMatch.AltName))
                {
                    firstMatch.Matched = true;
                    return firstMatch;
                }
            }

            target.SKU = target.SKU ?? target.AltName;

            // Name and altname same [Manualy add cost]
            var manualyMatched = set.Where(s =>
            !string.IsNullOrEmpty(s.AltName)
            && target.Name.Equals(s.Name, StringComparison.InvariantCultureIgnoreCase)
            && target.SKU.Equals(s.AltName, StringComparison.InvariantCultureIgnoreCase))
        .FirstOrDefault();

            if (manualyMatched != null)
            {
                manualyMatched.Matched = true;
                return manualyMatched;
            }
            

            // Duplicate matched
            var targetSize = Regex.Match(target.SKU, @"\d+").Value;
            if (string.IsNullOrEmpty(targetSize) && matched.Count > 1)
            {
                var dupMatched = matched.OrderBy(s => s.Price).First();
                dupMatched.Matched = true;
                return dupMatched;
            }

            // Multiple matched [Get sizing]
            if (!string.IsNullOrEmpty(targetSize) && matched.Count > 1)
            {

                foreach (var alt in matched.Where(s => s.AltName != null))
                {
                    alt.AltName = Regex.Match(alt.AltName, @"\d+").Value;
                }

                var sizeMatched = matched.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(targetSize.ToLower(), s.AltName.ToLower())).First();
                sizeMatched.Matched = true;
                return sizeMatched;
            }

            target.Matched = false;
            return target;
        }
    }
}
