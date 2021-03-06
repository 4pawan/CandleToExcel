using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Zerodha.Excel
{
    public class Candles
    {
        public DateTime Date { get; set; }
        public string DateFormated { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }
        public long Volume { get; set; }
        public double LowToHigh { get; set; }
        public double UpperTail { get; set; }
        public double LowerTail { get; set; }
        public double CENTHigh { get; set; } // high - open
        public double CENTLow { get; set; }  // open- low
        public double CENTClose { get; set; }  // open -close
        public double CENTLowToHigh { get; set; }
        public double Gap { get; set; }
        public bool IsLowerTailLarger { get; set; }
        public double CandleWeight { get; set; }

    }

    public class Response
    {
        public string status { get; set; }
        public Data data { get; set; }
    }
    public class Data
    {
        public List<List<object>> candles { get; set; }
    }

}
