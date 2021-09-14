using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Zerodha.Excel
{
    public class Excelhelper
    {
        public static void ExportToExcel()
        {
            string json = ReadJson();
            List<Candles> candleList = FormatJsonToObject(json);
            DataTable dt = ObjectToDataTable(candleList);
            CreateExcel(dt);
        }

        static string ReadJson()
        {
            string path = @"C:\\Project\\Kite.Exce\\Zerodha.Excel\\Zerodha.Excel\\input\\monthy.json";
            return File.ReadAllText(path);
        }

        static void CreateExcel(DataTable table)
        {
            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");

                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (string col in columns)
                    {
                        if (cellIndex == 0)
                        {
                            row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        }
                        else if (cellIndex == 4)
                        {
                            row.CreateCell(cellIndex).SetCellValue(long.Parse(dsrow[col].ToString()));
                        }
                        else
                        {
                            row.CreateCell(cellIndex).SetCellValue(Convert.ToDouble(dsrow[col]));
                        }

                        cellIndex++;
                    }

                    rowIndex++;
                }
                workbook.Write(fs);
            }

        }

        static List<Candles> FormatJsonToObject(string json)
        {
            var data = JsonConvert.DeserializeObject<Response>(json);
            List<Candles> candleList = new List<Candles>();

            // get all formating done with all calculations
            foreach (List<object> c in data.data.candles)
            {
                var candle = new Candles();
                candle.DateFormated = c[0].ToString();
                candle.Open = Convert.ToDouble(c[1]);
                candle.High = Convert.ToDouble(c[2]);
                candle.Low = Convert.ToDouble(c[3]);
                candle.Close = Convert.ToDouble(c[4]);
                candle.Volume = long.Parse(c[5].ToString());
                candle._CENTHigh = Convert.ToDouble(c[2]);
                candle._CENTLow = Convert.ToDouble(c[3]);
                candle._CENTClose = Convert.ToDouble(c[4]);
                candleList.Add(candle);
            }

            return candleList;
        }

        static DataTable ObjectToDataTable(List<Candles> candleList)
        {
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(candleList), (typeof(DataTable)));
            return table;
        }
    }
}
