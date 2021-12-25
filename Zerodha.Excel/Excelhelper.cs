using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Zerodha.Excel
{
    public class Excelhelper
    {
        private static string dateFormat = "dd-MM-yyyy:hh:mm";
        public static void ExportToExcel(string key)
        {
            string json = ReadJson();
            List<Candles> candleList = FormatJsonToObject(json);

            if (key == "M")
            {
                candleList = GetMonthlyData(candleList);
            }

            DataTable dt = ObjectToDataTable(candleList.OrderByDescending(c => c.Date).ToList());
            dt.Columns.Remove("Date");
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
                excelSheet.CreateFreezePane(0, 1);
                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    if (column.ColumnName == "IsLowerTailLarger")
                        continue;

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
                            DateTime date = DateTime.ParseExact(dsrow[col].ToString(), dateFormat, CultureInfo.InvariantCulture);
                            var cell = row.CreateCell(cellIndex);
                            if (IsMonday(date))
                            {
                                ICellStyle backGroundColorStyle = workbook.CreateCellStyle();
                                short colorBlue = HSSFColor.Blue.Index;
                                backGroundColorStyle.FillForegroundColor = colorBlue;
                                backGroundColorStyle.FillPattern = FillPattern.LessDots;
                                cell.CellStyle = backGroundColorStyle;
                            }
                            cell.SetCellValue(dsrow[col].ToString());
                        }
                        else if (cellIndex == 5)
                        {
                            row.CreateCell(cellIndex).SetCellValue(long.Parse(dsrow[col].ToString()));
                        }
                        else if (cellIndex == 8)
                        {
                            ICellStyle backGroundColorStyle = workbook.CreateCellStyle();
                            short colorBlue = HSSFColor.Red.Index;
                            backGroundColorStyle.FillForegroundColor = colorBlue;
                            backGroundColorStyle.FillPattern = FillPattern.LessDots;
                            var cell = row.CreateCell(cellIndex);
                            var rowVal = Convert.ToDouble(dsrow[col]);
                            bool isLowerTailLarger = Convert.ToBoolean(dsrow[14]);
                            cell.SetCellValue(rowVal);
                            if (isLowerTailLarger)
                            {
                                cell.CellStyle = backGroundColorStyle;
                            }
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
                var _date = DateTime.Parse(Convert.ToString(c[0]));
                candle.Date = _date;
                double Open = Convert.ToDouble(c[1]);
                double High = Convert.ToDouble(c[2]);
                double Low = Convert.ToDouble(c[3]);
                double Close = Convert.ToDouble(c[4]);
                double LowToHigh = High - Low;
                candle.DateFormated = _date.ToString(dateFormat);
                candle.Open = Open;
                candle.High = High;
                candle.Low = Low;
                candle.Close = Close;
                candle.Volume = long.Parse(c[5].ToString());
                candle.LowToHigh = LowToHigh;
                candle.Gap = candleList.Any() ? Open - candleList.Last().Close : 0;
                candle.CENTHigh = ((High - Open) / Open) * 100;
                candle.CENTLow = ((Open - Low) / Open) * 100;
                candle.CENTClose = ((Close - Open) / Open) * 100;
                candle.CENTLowToHigh = (LowToHigh / Low) * 100;
                candle.CandleWeight = Close - Open;
                SetTailProperty(candle);
                candleList.Add(candle);
            }

            return candleList;
        }

        static DataTable ObjectToDataTable(List<Candles> candleList)
        {
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(candleList), (typeof(DataTable)));
            return table;
        }
        static List<Candles> GetMonthlyData(List<Candles> candleList)
        {
            List<Candles> monthlyList = new List<Candles>();
            DateTime mindate = candleList.Min(d => d.Date);
            DateTime maxdate = candleList.Max(d => d.Date);

            for (int i = mindate.Year; i <= maxdate.Year; i++)
            {
                int month = mindate.Year == i ? mindate.Month : 1;
                for (int m = month; m <= 12; m++)
                {
                    var rangeMonthlyDates = candleList.Where(c => c.Date.Year == i && c.Date.Month == m);
                    if (!rangeMonthlyDates.Any())
                        continue;

                    var candle = new Candles();
                    DateTime date = new DateTime(i, m, 1);
                    double open = rangeMonthlyDates.First().Open;
                    double high = rangeMonthlyDates.Max(c => c.High);
                    double low = rangeMonthlyDates.Min(c => c.Low);
                    double close = rangeMonthlyDates.Last().Close;
                    double lowToHigh = high - low;
                    candle.Date = date;
                    candle.DateFormated = date.ToString(dateFormat);
                    candle.Open = open;
                    candle.Low = low;
                    candle.High = high;
                    candle.Close = close;
                    candle.LowToHigh = lowToHigh;
                    candle.CENTHigh = ((high - open) / open) * 100;
                    candle.CENTLow = ((open - low) / open) * 100;
                    candle.CENTClose = ((close - open) / open) * 100;
                    candle.CENTLowToHigh = (lowToHigh / low) * 100;
                    monthlyList.Add(candle);
                }

            }
            return monthlyList;
        }
        static void SetTailProperty(Candles candle)
        {
            if (candle.Close == candle.Open)
            {
                candle.UpperTail = 0;
                candle.LowerTail = 0;
            }

            if (candle.Close > candle.Open)
            {
                //when green candle
                candle.UpperTail = candle.High - candle.Close;
                candle.LowerTail = candle.Open - candle.Low;
            }
            else
            {
                //when red candle
                candle.UpperTail = candle.High - candle.Open;
                candle.LowerTail = candle.Close - candle.Low;
            }
            candle.IsLowerTailLarger = candle.LowerTail > candle.UpperTail;
        }

        static bool IsMonday(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Monday;
        }
    }
}
