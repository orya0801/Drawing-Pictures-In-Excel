using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDrawing
{
    class Excel
    {
        _Application excel = new _Excel.Application();
        Workbook xlWorkBook;
        Worksheet xlWorkSheet;
        _Excel.Range xlRange;


        Picture pict;
        Color[,] colors;
        bool[,] usedCells;


        public string Path { get; set; } = "";
        public int Sheet { get; set; }

        public Excel(){}

        public Excel(string path, int sheet)
        {
            excel.DisplayAlerts = false;
            Path = path;
            Sheet = sheet;
            xlWorkBook = excel.Workbooks.Open(Path);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[Sheet];
            xlRange = xlWorkSheet.UsedRange;
        }

        private void SetupExcelList(int height, int width)
        {
            string endColumn = CalculateRange(width);
            xlRange = xlWorkSheet.Range[$"A1:{endColumn}{height}"];

            // Задание ширины в символах 
            xlRange.ColumnWidth = 0.125;
            // Задание высоты в пунктах (пикселах)
            xlRange.RowHeight = 2;
            // Настройка цвета - установка прозрачного цвета
            //xlRange.Borders.Color = Color.Transparent;
        }

        public string CalculateRange(int width)
        {
            string range = "";
            List<int> rangeInNumbers = new List<int>();

            while(width > 0)
            {
                rangeInNumbers.Add(width % 26);
                width /= 26;
            }

            for (int i = 0; i < rangeInNumbers.Count; i++)
            {
                switch (rangeInNumbers[i])
                {
                    case 1:
                        range += "A";
                        break;
                    case 2:
                        range += "B";
                        break;
                    case 3:
                        range += "C";
                        break;
                    case 4:
                        range += "D";
                        break;
                    case 5:
                        range += "E";
                        break;
                    case 6:
                        range += "F";
                        break;
                    case 7:
                        range += "G";
                        break;
                    case 8:
                        range += "H";
                        break;
                    case 9:
                        range += "I";
                        break;
                    case 10:
                        range += "J";
                        break;
                    case 11:
                        range += "K";
                        break;
                    case 12:
                        range += "L";
                        break;
                    case 13:
                        range += "M";
                        break;
                    case 14:
                        range += "N";
                        break;
                    case 15:
                        range += "O";
                        break;
                    case 16:
                        range += "P";
                        break;
                    case 17:
                        range += "Q";
                        break;
                    case 18:
                        range += "R";
                        break;
                    case 19:
                        range += "S";
                        break;
                    case 20:
                        range += "T";
                        break;
                    case 21:
                        range += "U";
                        break;
                    case 22:
                        range += "V";
                        break;
                    case 23:
                        range += "W";
                        break;
                    case 24:
                        range += "X";
                        break;
                    case 25:
                        range += "Y";
                        break;
                    default:
                        range += "Z";
                        rangeInNumbers[i + 1]--;
                        break;
                }
            }

            char[] range_arr = range.ToCharArray();
            Array.Reverse(range_arr);
            Console.WriteLine(new string(range_arr));
            return new string(range_arr);
        }

        public void DrawPicture(string path)
        {
            pict = new Picture(path);
            colors = pict.GetColors();
            //usedCells = new bool[colors.GetLength(0), colors.GetLength(1)];
            SetupExcelList(pict.Height, pict.Width);
            try
            {
                Parallel.For(1, xlRange.Rows.Count + 1, y =>
                {
                    for (int x = 1; x <= xlRange.Columns.Count; x++)
                    {
                        var cell = xlRange.Cells[y, x];
                        cell.Interior.Color = colors[y - 1, x - 1];
                        //cell.Borders.Color = colors[y - 1, x - 1];  
                    }
                    Console.WriteLine($"Ряд {y} нарисован...");
                });
                Thread.Sleep(3000);
                // DrawingProcess(xlRange.Rows.Count);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.InnerException.Message);
                Save();
                Close();
                QuitExcel();
                Environment.Exit(1);
            }    
        }

        private void DrawingProcess(int rows)
        {
            for (int y = 1; y <= rows; y++)
            {
                for (int x = 1; x <= xlRange.Columns.Count; x++)
                {
                        var cell = xlRange.Cells[y, x];
                        cell.Interior.Color = colors[y - 1, x - 1];
                        cell.Borders.Color = colors[y - 1, x - 1];
                        //usedCells[y - 1, x - 1] = true;  
                }
                Console.WriteLine($"Ряд {y} нарисован...");
            }
            //Thread.Sleep(1000);
        }

        public void CreateNewFile()
        {
            this.xlWorkBook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.xlWorkSheet = xlWorkBook.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            Worksheet newSheet = xlWorkBook.Worksheets.Add(After: xlWorkSheet);
        }

        public void Save()
        {
            xlWorkBook.Save();
        }

        public void SaveAs(string path)
        {
            xlWorkBook.SaveAs(path);
        }

        public void Close()
        {
            xlWorkBook.Close();
        }

        public void QuitExcel()
        {
            excel.Quit();
        }

        public void OpenExcel()
        {
            Process.Start(Path);
        }
    }
}

