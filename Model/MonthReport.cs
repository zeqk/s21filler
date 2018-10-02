using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace S21Filler.Model
{
    public class MonthReport
    {
        public int Month { get; set; }
        public int Placements { get; set; }
        public int VideoShowings { get; set; }
        public int Hours { get; set; }
        public int ReturnVisits { get; set; }
        public int Studies { get; set; }
        public string Remarks { get; set; }
    }
}
