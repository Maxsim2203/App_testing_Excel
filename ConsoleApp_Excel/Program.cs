// See https://aka.ms/new-console-template for more information
using System;
using System.Linq;
internal class Program
{
    private static void Main(string[] args)
    {
        //using DocumentFormat.OpenXml;
        //using DocumentFormat.OpenXml.Packaging;
        //using DocumentFormat.OpenXml.Spreadsheet;
        bool ind = true;
        while (ind)
        {
            Console.WriteLine("Выберите действие (стрелками переместите курсор на нужный пункт меню и нажмите ENTER):");
            int top = Console.CursorTop;
            int y = top;

            Console.WriteLine("1. Указать путь к файлу с данными");
            Console.WriteLine("2. Указать наименование товара и вывести список клентов, заказавших товар");
            Console.WriteLine("3. Определить золотого клиента");

            int down = Console.CursorTop;

            Console.CursorSize = 100;
            Console.CursorTop = top;

            ConsoleKey key;
            while ((key = Console.ReadKey(true).Key) != ConsoleKey.Enter)
            {
                if (key == ConsoleKey.UpArrow)
                {
                    if (y > top)
                    {
                        y--;
                        Console.CursorTop = y;
                    }
                }
                else if (key == ConsoleKey.DownArrow)
                {
                    if (y < down - 1)
                    {
                        y++;
                        Console.CursorTop = y;
                    }
                }
            }

            Console.CursorTop = down;
            
                if (y == top)
                {
                    Console.Write("");
                    Console.Write("Введите полный путь к файлу: ");
                    string? path_name = Console.ReadLine();
                    bool exist = File.Exists(path_name);
                    if (exist == true)
                    {
                        Console.Write("Файл существует, для продолжения работы нажмите ENTER...ESC - конец работы");
                    }
                    if (exist == false)
                    {
                        Console.Write("Путь к файлу указан неверно, для продолжения работы нажмите ENTER...ESC - конец работы");
                    }
                else if (y == top + 1)
                {
                    Console.WriteLine("два");
                }
                else if (y == top + 2)
                {
                    Console.WriteLine("три");
                }
                }
            while (Console.ReadKey().Key != ConsoleKey.Enter) { } // || (Console.ReadKey().Key != ConsoleKey.Escape)) { }
            Console.Clear();
           
            //if (Console.ReadKey().Key != ConsoleKey.Escape)
              //  ind = false;
            //return;
        }
    }
}


/*string pathToFile = @"D:\data.xlsx";
//Console.WriteLine("Hello, World!");
//Создаём приложение.
Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
//Открываем книгу.                                                                                                                                                        
Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
//Выбираем таблицу(лист).
Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Товары;

// Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
int numCol = 2;

Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

// Выходим из программы Excel.
ObjExcel.Quit();
*/