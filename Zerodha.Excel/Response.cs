using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zerodha.Excel
{
    public class Candles
    {
        public string DateFormated { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }
        public long Volume { get; set; }
        public double _CENTHigh { get; set; } // high - open
        public double _CENTLow { get; set; }  // open- low
        public double _CENTClose { get; set; }  // open -close
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
