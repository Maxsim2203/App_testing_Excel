// See https://aka.ms/new-console-template for more information
using System;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
//using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    private static void Main(string[] args)
    {
        //using DocumentFormat.OpenXml;
        //using DocumentFormat.OpenXml.Packaging;
        //using DocumentFormat.OpenXml.Spreadsheet;
        bool ind = true;
        while (ind = true)
        {
            //Делаем простой вывод на экран
            Console.WriteLine("Выберите действие (стрелками переместите курсор на нужный пункт меню и нажмите ENTER):");
            int top = Console.CursorTop;
            int y = top;

            Console.WriteLine("1. Указать путь к файлу с данными");
            Console.WriteLine("2. Указать наименование товара и вывести список клентов, заказавших товар");
            Console.WriteLine("3. Изменить контактное лицо клиена по критериям");
            Console.WriteLine("4. Определить золотого клиента");

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
                    //Console.Write("");
                    Console.Write("Введите полный путь к файлу: ");
                    string? path_file = Console.ReadLine();
                    bool exist = File.Exists(path_file);
                    if (exist == true)
                    {
                        Console.Clear();
                        Console.WriteLine("Файл существует.");
                    }
                    if (exist == false)
                    {
                        Console.Clear();
                        Console.WriteLine("Путь к файлу указан неверно!");
                    }
                       
                else if (y == top + 1)
                    {
                    if (exist == true)
                    {
                      Console.Clear();
                      Console.WriteLine("Укажите наименование товара");
                    }
                    if (exist == false)
                    {
                        Console.Clear();
                        Console.WriteLine("Товар не выбран!");
                    }
                  Console.Clear();
                   
                }
                else if (y == top + 2)
                {
                    //string pathToFile = "D:\\1.xlsx";
                    ////Открываем книгу.                                                                                                                                                        
                    //Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    ////Выбираем таблицу(лист).
                    //Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                    //ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Товары;

                    //// Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
                    //int numCol = 2;

                    //Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
                    //System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
                    //string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

                    //// Выходим из программы Excel.
                    //ObjExcel.Quit();


                }
                else if (y == top + 3)
                {
                    Console.WriteLine("четыре");
                }
            }
                 
            Console.WriteLine("ENTER - продолжение работы");
            Console.WriteLine("ESC - выход");
            key = Console.ReadKey().Key;
            while ((key != ConsoleKey.Enter) & (key != ConsoleKey.Escape)) { }
            Console.Clear();
            if (key == ConsoleKey.Escape)
            {
               Console.WriteLine("ППрограмма завершила работу. До свидания."); //Два П т.к по нажатию клавиши ESC убирается первый символ 
               break;
            }
        }
    }
}


//string pathToFile = @"D:\data.xlsx";
////Console.WriteLine("Hello, World!");
////Создаём приложение.
//Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
////Открываем книгу.                                                                                                                                                        
//Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
////Выбираем таблицу(лист).
//Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
//ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Товары;

//// Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
//int numCol = 2;

//Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
//System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
//string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

//// Выходим из программы Excel.
//ObjExcel.Quit();
