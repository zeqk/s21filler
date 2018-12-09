using System;
using System.Collections.Generic;

namespace S21Filler.Model
{
    public class YearRecord
    {

        public YearRecord()
        {
            Reports = new List<MonthReport>();
        }

        public string Name { get; set; }
        public string Number { get; set; }
        public Genders Gender { get; set; }
        public string HomeAddress { get; set; }
        public string HomeTelephone { get; set; }
        public string MobileTelephone { get; set; }
        public DateTime DateOfBirth { get; set; }
        public DateTime? ImmersedDate { get; set; }
        public string Anointed { get; set; }
        public bool E { get; set; }
        public bool MS { get; set; }
        public bool RP { get; set; }
        public int Year { get; set; }
        public IList<MonthReport> Reports { get; set; }
        public MonthReport Totals { get; set; }
    }
}
