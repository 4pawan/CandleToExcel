using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Zerodha.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Directory.GetCurrentDirectory();
            Console.WriteLine(path + "\n\n");
            Console.WriteLine("Press key M to get monthly data");
            Console.WriteLine("Press key D to get weekly data");
            Console.WriteLine("Press key W to get daily data");
            var input = Console.ReadKey();
            
            Excelhelper.ExportToExcel(input.Key.ToString());
            Console.WriteLine("------Press any key to exit! -----------------");
            Console.ReadKey();
        }
    }
}
