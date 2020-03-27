using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ShopHelper.Models;

namespace ShopHelper
{
    public static class MatchingHelper
    {
        public static Item Match(Item target, IEnumerable<Item> set)
        {
            var matched =
                set.Where(s => target.Name.GetName() == s.Name.GetName()).ToList();

            var matched_targetLonger =
                set.Where(s => target.Name.GetName().Contains(s.Name.GetName())).ToList();

            var matched_setLonger =
                set.Where(s => s.Name.GetName().Contains(target.Name.GetName())).ToList();

            if (!matched.Any() && matched_targetLonger.Any() &&
                (matched_targetLonger.Count > matched.Count))
            {
                matched = matched_targetLonger;
            }

            if (!matched.Any() && matched_setLonger.Any() &&
                (matched_setLonger.Count > matched.Count))
            {
                matched = matched_setLonger;
            }

            if (!matched.Any())
            {
                // 90% matched
                matched = set.Where(s => CompareHelper.Compare(s.Name.GetName(), target.Name.GetName()) < target.Name.Length * 0.1).ToList();
            }

            if (!matched.Any())
            {
                // 80% matched
                matched = set.Where(s => CompareHelper.Compare(s.Name.GetName(), target.Name.GetName()) < target.Name.Length * 0.2).ToList();
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

            target.SKU = target.SKU ?? target.AltName ?? "";

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
            var targetSize = target.SKU;
            if (string.IsNullOrEmpty(targetSize) && matched.Any())
            {
                var dupMatched = matched.OrderBy(s => s.Price).First();
                dupMatched.Matched = true;
                return dupMatched;
            }

            // Multiple matched [Get sizing]
            if (!string.IsNullOrEmpty(targetSize) && matched.Any())
            {
                // Duplicate matched but target has sku

                if (matched.All(x => string.IsNullOrEmpty(x.AltName)))
                {
                    var moreMath = matched.OrderByDescending(x => x.Stock).First();
                    moreMath.Matched = true;
                    return moreMath;
                }

                foreach (var alt in matched.Where(s => s.AltName != null))
                {
                    alt.AltName = alt.AltName;
                }
                
                var sizeMatched = matched.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(targetSize.GetName(), s.AltName.GetName())).First();

                // Price so differance
                var targetPrice = target.Price;
                var matchedPrice = sizeMatched.Price;

                if (Math.Abs(targetPrice - matchedPrice) > 100)
                {
                    sizeMatched = CompareHelper.GetClosedPrice(targetPrice, matched);
                }

                sizeMatched.MultiStocks = string.Join(", ", matched.Select(x => x.AltName + " : " + x.Stock));
                sizeMatched.MultiPrices = string.Join(", ", matched.Select(x => x.AltName + " : " + x.Price));

                sizeMatched.Matched = true;
                return sizeMatched;
            }

            target.Matched = false;
            return target;
        }

        public static Item Match_Stable( Item target, IEnumerable<Item> set)
        {
            var matched =
                set.Where(s => target.Name.GetName() == s.Name.GetName()).ToList();

            var matched_targetLonger =
                set.Where(s => target.Name.GetName().Contains(s.Name.GetName())).ToList();

            var matched_setLonger =
                set.Where(s => s.Name.GetName().Contains(target.Name.GetName())).ToList();

            if(!matched.Any() && matched_targetLonger.Any() &&
                (matched_targetLonger.Count > matched.Count))
            {
                matched = matched_targetLonger;
            }

            if (!matched.Any() && matched_setLonger.Any() &&
                (matched_setLonger.Count > matched.Count))
            {
                matched = matched_setLonger;
            }

            if (!matched.Any())
            {
                // 90% matched
                matched = set.Where(s => CompareHelper.Compare(s.Name.GetName(), target.Name.GetName()) < target.Name.Length * 0.1).ToList();
            }

            if (!matched.Any())
            {
                // 80% matched
                matched = set.Where(s => CompareHelper.Compare(s.Name.GetName(), target.Name.GetName()) < target.Name.Length * 0.2).ToList();
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

            target.SKU = target.SKU ?? target.AltName ?? "";

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
            var targetSize = string.IsNullOrEmpty(target.SKU) ? null : Regex.Match(target.SKU, @"\d+").Value;
            if (string.IsNullOrEmpty(targetSize) && matched.Any())
            {
                var dupMatched = matched.OrderBy(s => s.Price).First();
                dupMatched.Matched = true;
                return dupMatched;
            }

            // Multiple matched [Get sizing]
            if (!string.IsNullOrEmpty(targetSize) && matched.Any())
            {
                // Duplicate matched but target has sku

                if(matched.All(x => string.IsNullOrEmpty(x.AltName)))
                {
                    var moreMath = matched.OrderByDescending(x => x.Stock).First();
                    moreMath.Matched = true;
                    return moreMath;
                }

                foreach (var alt in matched.Where(s => s.AltName != null))
                {
                    alt.AltName = Regex.Match(alt.AltName, @"\d+").Value;
                }

                var sizeMatched = matched.Where(s => s.AltName != null).OrderBy(s => CompareHelper.Compare(targetSize.GetName(), s.AltName.GetName())).First();
                sizeMatched.Matched = true;
                return sizeMatched;
            }

            target.Matched = false;
            return target;
        }

        public static Item MatchSku(Item target, IEnumerable<Item> set)
        {
            var matched = set.FirstOrDefault(s => target.SKU.GetName() == s.SKU.GetName());

            if(matched != null)
            {
                matched.Matched = true;
                return matched;
            }

            return target;
        }

        private static string GetName(this string name)
        {
            var startPoint = name.IndexOf("[ ");
            var endPoint = name.IndexOf(" ]");

            if (startPoint > -1 && endPoint > -1)
            {
                int length = endPoint - startPoint -2;

                return name.Substring(startPoint + 2, length).Trim().GetName();
            }
            return name;
        }
    }
}
