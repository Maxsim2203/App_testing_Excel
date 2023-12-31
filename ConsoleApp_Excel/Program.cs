﻿// Образец указания пути к файлу: D:\1.xlsx
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
using Microsoft.Office.Interop.Excel;

internal class Program
{
    private static void Main(string[] args)
    {
        bool ind = true;
        string? path_file = " ";
        string? kod_tov = " ";
        string? kol_tov = " ";
        string? kod_cli = " ";
        string? naim_cli = " ";
        string? cost_tov = " ";
        string? date_zak = " ";
        string? FIO_cli = " ";
        string? s = " ";
        int Max = 0;
        string? naim_cli_max = " ";
        bool check_1 = false; //Контроль наличия товара
        bool check_2 = false; //Контроль наличия заказа
        
        while (ind == true)
        {
            //Формируем меню программы
            Console.WriteLine("Выберите действие (стрелками переместите курсор на нужный пункт меню и нажмите ENTER):");
            int top = Console.CursorTop;
            int y = top;

            Console.WriteLine("1. Указать путь к файлу с данными");
            Console.WriteLine("2. Указать наименование товара и вывести список клиентов, заказавших товар");
            Console.WriteLine("3. Изменить контактное лицо клиента");
            Console.WriteLine("4. Определить золотого клиента");
            Console.WriteLine("5. Выход из программы");

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

            if (y == top) // Если выбран 1-й пункт меню
            {
                Console.Write("Введите полный путь к файлу: ");
                path_file = Console.ReadLine();
                bool exist = File.Exists(path_file);
                if (exist == true)
                {
                    Console.Clear();
                    Console.WriteLine("Файл существует. Можете продолжить работу.");
                }
                if (exist == false)
                {
                    Console.Clear();
                    Console.WriteLine("Путь к файлу указан неверно!");
                }
            }
            else if (y == top + 1) // Если выбран 2-й пункт меню
            {
                Console.Clear();

                if (path_file != " ") //Проверяем - указан ли путь к файлу (пункт 1 программы)
                {
                    Console.WriteLine("Укажите наименование товара: ");
                    string? nam_tov = Console.ReadLine();
                    XLWorkbook workbook;
                    using (workbook = new XLWorkbook(@path_file))
                    {
                        // Получение 1 листа из книги
                        var worksheet_1 = workbook.Worksheet(1);

                        // Определение первой и последней строки в 1 листе
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
                                kod_cli = row_3.Cell("C").Value.ToString();
                                kol_tov = row_3.Cell("E").Value.ToString();
                                date_zak = row_3.Cell("F").Value.ToString();

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

                                Console.WriteLine(naim_cli + " " + kol_tov + " " + summ_zak + " " + date_zak);

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
                if (path_file == " ") //Если не указан путь к файлу (не выполнен пункт 1 программы)
                {
                    Console.WriteLine("Не указан путь к файлу!");
                }
            }

            else if (y == top + 2) // Если выбран 3-й пункт меню
            {
                Console.Clear();

                if (path_file != " ") //Если указан путь к файлу (выполнен пункт 1 программы)
                {
                    Console.WriteLine("Укажите наименование огранизации: ");
                    string? nai_cli = Console.ReadLine();
                    XLWorkbook workbook;
                    using (workbook = new XLWorkbook(@path_file))
                    {
                        // Получение второго листа из книги
                        var worksheet_2 = workbook.Worksheet(2);
                        // Определение первой и последней строки в листе
                        var firstRow_2 = worksheet_2.FirstRowUsed();
                        var lastRow_2 = worksheet_2.LastRowUsed();
                        // Находим на 2 листе ячейку с наименованием организации
                        foreach (var row in worksheet_2.Rows(firstRow_2.RowNumber(), lastRow_2.RowNumber()))
                        {
                            if (row.Cell("B").Value.ToString() == nai_cli)
                            {
                                check_2 = true;
                                naim_cli = row.Cell("B").Value.ToString();
                                FIO_cli = row.Cell("D").Value.ToString();
                                Console.WriteLine(FIO_cli);
                                Console.WriteLine("Укажите ФИО нового контактного лица");
                                string? FIO_new_cli = Console.ReadLine();
                                row.Cell("D").Value = FIO_new_cli;
                                workbook.Save();
                            }
                        }
                        if (check_2 == false)
                        {
                            Console.WriteLine("Организация не найдена. Повторите попытку.");
                        }
                        else
                        {
                            Console.WriteLine("Изменения внесены в таблицу");
                        }
                    }
                }
                if (path_file == " ") //Если не указан путь к файлу (не выполнен пункт 1 программы)
                {
                    Console.WriteLine("Не указан путь к файлу!");
                }
            }

            else if (y == top + 3) // Если выбран 4-й пункт меню 
            {
                if (path_file != " ") //Если указан путь к файлу (выполнен пункт 1 программы)
                {
                    XLWorkbook workbook;
                    using (workbook = new XLWorkbook(@path_file))
                    {
                        // Получение первого, второго  и третьего листа из книги
                        var worksheet_2 = workbook.Worksheet(2);
                        var worksheet_3 = workbook.Worksheet(3);
                        var worksheet_1 = workbook.Worksheet(1);
                        // Определение первой и последней строки во втором листе
                        var firstRow_2 = worksheet_2.FirstRowUsed();
                        var lastRow_2 = worksheet_2.LastRowUsed();
                        // Определение первой и последней строки в третьем листе
                        var firstRow_3 = worksheet_3.FirstRowUsed();
                        var lastRow_3 = worksheet_3.LastRowUsed();
                        // Определение первой и последней строки в первом листе
                        var firstRow_1 = worksheet_1.FirstRowUsed();
                        var lastRow_1 = worksheet_1.LastRowUsed();

                        // Находим на 2 листе ячейку с наименованием организации
                        foreach (var row_2 in worksheet_2.Rows(firstRow_2.RowNumber() + 1, lastRow_2.RowNumber()))
                        {
                            row_2.Cell("E").Value = 0;
                            workbook.Save();
                            kod_cli = row_2.Cell("A").Value.ToString();
                            foreach (var row_3 in worksheet_3.Rows(firstRow_3.RowNumber() + 1, lastRow_3.RowNumber()))
                            {
                                if (row_3.Cell("C").Value.ToString() == kod_cli)
                                {
                                    kol_tov = row_3.Cell("E").Value.ToString();
                                    kod_tov = row_3.Cell("B").Value.ToString();
                                    foreach (var row_1 in worksheet_1.Rows(firstRow_1.RowNumber() + 1, lastRow_1.RowNumber()))
                                    {
                                        if (row_1.Cell("A").Value.ToString() == kod_tov)
                                        {
                                            cost_tov = row_1.Cell("D").Value.ToString();
                                        }
                                    }
                                    int kol_tov_ = Int32.Parse(kol_tov);
                                    int cost_tov_ = Int32.Parse(cost_tov);
                                    int summ_zak_ = kol_tov_ * cost_tov_;
                                    s = row_2.Cell("E").Value.ToString();
                                    int ss = Int32.Parse(s);
                                    row_2.Cell("E").Value = ss + summ_zak_;
                                }
                            }
                        }
                        workbook.Save();
                        foreach (var row_2 in worksheet_2.Rows(firstRow_2.RowNumber() + 1, lastRow_2.RowNumber()))
                        {
                            naim_cli = row_2.Cell("B").Value.ToString();
                            string a = row_2.Cell("E").Value.ToString();
                            int a_ = Int32.Parse(a);
                            if (a_ > Max)
                            {
                                naim_cli_max = naim_cli;
                                Max = a_;
                            }
                        }
                        Console.Clear();
                        Console.WriteLine("Золотой клиент: " + naim_cli_max);
                    }
                }
                if (path_file == " ") //Если не указан путь к файлу (не выполнен пункт 1 программы)
                {
                    Console.Clear();
                    Console.WriteLine("Не указан путь к файлу!");
                }

            }
            else if (y == top + 4) // Если выбран 5-й пункт меню
            {
                Console.Clear();
                Console.WriteLine("Программа завершила работу. До свидания.");
                break;
            }

            //Очистка переменных
            kod_tov = " ";
            kol_tov = " ";
            kod_cli = " ";
            naim_cli = " ";
            cost_tov = " ";
            date_zak = " ";
            check_1 = false;
            check_2 = false;
            Max = 0;

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