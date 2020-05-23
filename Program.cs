using System;
using System.IO;
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
            try
            {
                var pathToExcel = Path.GetFullPath(@"test1.xlsx");
                Excel excel = new Excel(pathToExcel, 1);

                try
                {
                    Console.WriteLine("Введите путь до картинки:");
                    string pathToImage = Console.ReadLine();
                    excel.DrawPicture(pathToImage);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.InnerException.Message);
                }
                finally
                {
                    excel.Save();
                    excel.Close();
                    excel.QuitExcel();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
