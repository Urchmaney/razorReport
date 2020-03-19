using System.Collections.Generic;

namespace razorReport.Models
{
    public class NamedCashDenomnation
    {
        public string Name { get; set; }

        public Dictionary<string,double> CashDenomination { get; set; }
    }
}