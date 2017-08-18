using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCalc
{
    public class Employee
    {
        internal int SocialPoints;
        internal int TechPoints;

        public string Name;

        public int CountOfLead { get; internal set; }

        public List<Tuple<string, int>> ListSummaryPoints;
        public List<Tuple<string, float>> ListAveragePoints;
        public List<List<Tuple<string, int>>> ListPoints;
    }
}
