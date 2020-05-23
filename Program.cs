using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDrawing
{
    class Program
    {
        static void Main(string[] args)
        {
            var pathToExcel = System.IO.Path.GetFullPath(@"test1.xlsx");
            Excel excel = new Excel(pathToExcel, 1);
            Console.WriteLine("Введите путь до картинки:");
            string pathToImage = Console.ReadLine();
            excel.DrawPicture(pathToImage);
            excel.Save();
            excel.Close();
            excel.QuitExcel();

        }
    }
}
