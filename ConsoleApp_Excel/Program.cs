// See https://aka.ms/new-console-template for more information
using System;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using static ClosedXML.Excel.XLWorkbook;
using DocumentFormat.OpenXml.Drawing.Charts;

internal class Program
{
    private static void Main(string[] args)
    {
        //using DocumentFormat.OpenXml;
        //using DocumentFormat.OpenXml.Packaging;
        //using DocumentFormat.OpenXml.Spreadsheet;
        bool ind = true;
        string? path_file = " ";
        string? kod_tov = " ";
        string? kol_tov = " ";
        string? kod_cli = " ";
        string? naim_cli = " ";
        string? cost_tov = " ";
        string? date_zak = " ";
        bool check_1 = false; //Контроль наличия товара
        bool check_2 = false; //Контроль наличия заказа
        while (ind == true)
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
                Console.Write("Введите полный путь к файлу: ");
                path_file = Console.ReadLine();
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
            }
            else if (y == top + 1)
            {
                Console.Clear();
                Console.WriteLine("Укажите наименование товара: ");
                string? nam_tov = Console.ReadLine();
                XLWorkbook workbook;
                //path_file = '@' + '"' + path_file +'"';
                //Console.WriteLine(path_file);
                path_file = @"D:\1.xlsx";
                using (workbook = new XLWorkbook(path_file))
                {
                // Получение первого листа из книги
                var worksheet_1 = workbook.Worksheet(1);

                // Определение первой и последней строки в листе
                var firstRow_1 = worksheet_1.FirstRowUsed();
                var lastRow_1 = worksheet_1.LastRowUsed();

                // Находим на 1 листе ячейку с наименованием указанного товара и запоминаем код и стоимость товара
                        foreach (var row in worksheet_1.Rows(firstRow_1.RowNumber(), lastRow_1.RowNumber()))
                        {
                            if (row.Cell("B").Value.ToString() == nam_tov)
                            {
                                check_1 = true;
                                kod_tov = row.Cell("A").Value.ToString();
                                cost_tov = row.Cell("D").Value.ToString();
                            }
                        }

                // Находим на 3 листе ячейку с кодом указанного товара и запоминаем код клиента 
                var worksheet_3 = workbook.Worksheet(3);
                var firstRow_3 = worksheet_3.FirstRowUsed();
                var lastRow_3 = worksheet_3.LastRowUsed();
                        foreach (var row_3 in worksheet_3.Rows(firstRow_3.RowNumber(), lastRow_3.RowNumber()))
                        {
                            if (row_3.Cell("B").Value.ToString() == kod_tov)
                            {
                                check_2 = true;
                                kod_cli = row_3.Cell("C").Value.ToString();  //287 - Чай
                                kol_tov = row_3.Cell("E").Value.ToString();  //5 - для первого заказа 10 - для второго
                                date_zak = row_3.Cell("F").Value.ToString(); //14.03.2023 - для первого заказа 22.06.2023 - для второго
                                
                                // Находим на 2 листе ячейку с кодом клиента и  запоминаем наименование клиента
                                var worksheet_2 = workbook.Worksheet(2);
                                var firstRow_2 = worksheet_2.FirstRowUsed();
                                var lastRow_2 = worksheet_2.LastRowUsed();
                                foreach (var row_2 in worksheet_2.Rows(firstRow_2.RowNumber(), lastRow_2.RowNumber()))
                                {
                                    if (row_2.Cell("A").Value.ToString() == kod_cli)
                                    {
                                        naim_cli = row_2.Cell("B").Value.ToString();
                                    }
                                }

                                date_zak = date_zak.Substring(0, date_zak.Length - 8);
                                int kol_tov_ = Int32.Parse(kol_tov);
                                int cost_tov_ = Int32.Parse(cost_tov);
                                int summ_zak_ = kol_tov_ * cost_tov_;
                                string summ_zak = Convert.ToString(summ_zak_);
                                                      
                                Console.WriteLine(naim_cli + " " +  kol_tov + " " + summ_zak + " " + date_zak);
                                
                            }
                        }
                }
                if (check_1 == false)
                { 
                    Console.WriteLine("Товар не найден!");
                }
                if (check_2 == false)
                {
                    Console.WriteLine("Заказов нет!");
                }
            }
            else if (y == top + 2)
            {
                Console.Clear();
                Console.WriteLine("Укажите наименование огранизации");
               // Console.WriteLine("Введите новое контакнтое лицо");
                
                //string? FIO_cli;
                //naim_cli

            }
            else if (y == top + 3)
            {
                Console.WriteLine("Золотой клиент");
            }
            path_file = " ";
            kod_tov = " ";
            kol_tov = " ";
            kod_cli = " ";
            naim_cli = " ";
            cost_tov = " ";
            date_zak = " ";
            check_1 = false; 
            check_2 = false;

            Console.WriteLine("ENTER - продолжение работы");
            Console.WriteLine("ESC - выход");
            key = Console.ReadKey().Key;
            while ((key != ConsoleKey.Enter) & (key != ConsoleKey.Escape)) { }
            Console.Clear();
            if (key == ConsoleKey.Escape)
            {
               Console.WriteLine(y);
               Console.WriteLine("Программа завершила работу. До свидания."); 
               break;
            }
        }
    }
}



//Console.WriteLine(kod_cli);
//Console.WriteLine(naim_cli);
//Console.WriteLine(kol_tov);
//Console.WriteLine(date_zak);
//key = Console.ReadKey().Key;


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

//else if (y == top + 2)
//{
//string pathToFile = "D:\\1.xlsx";
//////Открываем книгу.                                                                                                                                                        
//Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
//////Выбираем таблицу(лист).
//Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
//ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook;

//// Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
//int numCol = 4;

//Range usedColumn = ObjWorkSheet.Column[numCol];
//System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
//string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

//////Выходим из программы Excel.
//ObjExcel.Quit();

////Создание экземпляра Workbook
//Workbook workbook = new Workbook();

////Получение первой рабочей страницы
///                            //Console.Write(row.Cell("A").Value + " ");
//Console.Write(row.Cell("B").Value + " ");
//Console.WriteLine(row.Cell("D").Value);
//Console.WriteLine(cell);

// Определение столбца, в котором будет производиться поиск
//var column = worksheet.Column("B");
//var columnCells = column.CellsUsed();
//var cell = columnCells.FirstOrDefault();

/*foreach (var row in worksheet.Rows(firstRow.RowNumber(), lastRow.RowNumber()))
{
    if (row.Cell("B").Value.ToString() == nam_tov)
    {
        Console.Write(row.Cell("A").Value + " ");
        Console.Write(row.Cell("B").Value + " ");
        Console.WriteLine(row.Cell("D").Value);
        //Console.WriteLine(cell);
    }

}

    /*var cell = columnCells.F;
    // Получение значения ячейки
    //var cell = columnCells.First(c => c.Value.ToString().Contains(nam_tov));
    //var cell = columnCells.(1);


    Console.Write(cellValue);
    Console.WriteLine(nam_tov);

    var currentCell = columnCells.Cast<Cell>().FirstOrDefault();
    // Обработка каждой строки
    //foreach (var cell in row.CellsUsed())


    foreach (var row in worksheet.Rows(firstRow.RowNumber(), lastRow.RowNumber()))
    {
        // Получение значения ячейки
        //var cell = columnCells.First(c => c.Value.ToString().Contains(nam_tov));
        //var cell = columnCells.(1);

        var cellValue = currentCell.Value.ToString();
        Console.Write(cellValue);
        Console.WriteLine(nam_tov);

        if (cellValue != nam_tov)
        {
            Console.Write(row.Cell("A").Value + " ");
            Console.Write(row.Cell("B").Value + " ");
            Console.WriteLine(row.Cell("D").Value);*/



