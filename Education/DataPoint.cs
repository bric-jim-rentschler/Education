using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Education
{
    class DataPoint
    {
        private String periodType;
        private String category;
        private DateTime date;
        private String value;
        private String neum;

        public DataPoint(String periodType, String category, DateTime date, String value, String neum)
        {
            this.periodType = periodType;
            this.category = category;
            this.date = date;
            this.value = value;
            this.neum = neum;
        }

        public string PeriodType { get => periodType; set => periodType = value; }
        public string Category { get => category; set => category = value; }
        public DateTime Date { get => date; set => date = value; }
        public string Value { get => value; set => this.value = value; }
        public string Neum { get => neum; set => neum = value; }
    }

}
