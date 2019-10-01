using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShopHelper.Models
{
    public class PathCombination
    {
        public string FirstPath { get; set; }
        public string SecoundPath { get; set; }
        public string ThirdPath { get; set; }
        public string OutPutPath { get; set; }
        
        public PathCombination(string firstPath, string secoundPath, string thirdPath, string outPutPath)
        {
            FirstPath = firstPath;
            SecoundPath = secoundPath;
            ThirdPath = thirdPath;
            OutPutPath = outPutPath;
        }
        public PathCombination(string firstPath, string secoundPath, string outPutPath)
        {
            FirstPath = firstPath;
            SecoundPath = secoundPath;
            OutPutPath = outPutPath;
        }
    }
}
