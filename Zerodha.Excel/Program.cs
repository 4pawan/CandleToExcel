﻿using System;
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
            Excelhelper.ExportToExcel();
            Console.WriteLine("------Press any key to exit! -----------------");
            Console.ReadKey();
        }
    }
}