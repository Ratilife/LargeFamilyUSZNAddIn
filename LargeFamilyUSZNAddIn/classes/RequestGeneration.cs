using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LargeFamilyUSZNAddIn.classes
{
    internal class RequestGeneration
    {
        public void GetQueryResults(string SQLQuery)
        {
            string movieFilePath;
            OleDbConnection cn = null;
            OleDbDataAdapter da = null;
            DataTable dt = null;
            Excel.Worksheet ws = null;
            Excel.Application excelAppCurrentBook = Globals.ThisAddIn.Application;
            Excel.Application xlWorkbook = Globals.ThisAddIn.Application;
            int rowCount, colCount;
            /* 
            Эта процедура выполняет SQL-запрос к Excel файлу и записывает результаты в новую таблицу Excel.
            */

            // Выход из процедуры, если запрос пустой
            if (string.IsNullOrEmpty(SQLQuery))
            {
                MessageBox.Show("Вы не ввели запрос", "Отсутствует строка запроса", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверка, существует ли файл Movies.xlsx в той же папке, что и исполняемый файл
            movieFilePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Movies.xlsx");

            if (!System.IO.File.Exists(movieFilePath))
            {
                MessageBox.Show("Не удалось найти Movies.xlsx", "Файл не найден", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Создание и открытие соединения с файлом Movies.xlsx
                cn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={movieFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';");
                cn.Open();

                // Создание и заполнение DataTable с помощью SQL-запроса
                da = new OleDbDataAdapter(SQLQuery, cn);
                dt = new DataTable();
                da.Fill(dt);

                // Подсчет количества строк, возвращенных запросом
                rowCount = dt.Rows.Count;

                Console.WriteLine($"{rowCount} строк(и), {SQLQuery}");

                // Выход из процедуры, если результат запроса пуст
                if (rowCount == 0)
                {
                    MessageBox.Show("Запрос не вернул результатов", "Нет результатов", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Подсчет количества колонок, возвращенных запросом
                colCount = dt.Columns.Count;

                // Создание нового листа Excel
                ws = xlWorkbook.Sheets.Add();
                ws.Name = "QueryResults";

                // Форматирование заголовков таблицы
                Excel.Range headerRange = ws.Range["A1"].Resize[1, colCount];
                headerRange.Interior.Color = Excel.XlRgbColor.rgbCornflowerBlue;  // Заливка цветом
                headerRange.Font.Color = Excel.XlRgbColor.rgbWhite;  // Цвет шрифта
                headerRange.Font.Bold = true;  // Полужирный шрифт

                // Запись имен колонок в первую строку листа Excel
                for (int i = 0; i < colCount; i++)
                {
                    ws.Cells[1, i + 1] = dt.Columns[i].ColumnName;

                    // Применение специального формата для колонок с датами
                    if (dt.Columns[i].DataType == typeof(DateTime))
                    {
                        Excel.Range dateRange = ws.Range[ws.Cells[2, i + 1], ws.Cells[rowCount + 1, i + 1]];
                        dateRange.NumberFormat = "dd mmm yyyy";  // Формат даты
                    }
                }

                // Копирование данных из DataTable на лист Excel
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        ws.Cells[i + 2, j + 1] = dt.Rows[i][j];
                    }
                }

                // Автоподбор ширины колонок
                ws.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // TODO написать обработки ошибок
            }
        }

        public void DeleteAllButMenuSheet()
        {
            //?
        }


        public void GetQueryResults_example(string SQLQuery)
        {
            string movieFilePath;
            OleDbConnection cn = null;
            OleDbDataAdapter da = null;
            DataTable dt = null;
            Excel.Worksheet ws = null;
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
            //текущая книга
            Excel.Application xlWorkbook = Globals.ThisAddIn.Application;
            int rowCount, colCount;

            /* 
            Эта процедура выполняет SQL-запрос к Excel файлу и записывает результаты в новую таблицу Excel.
            */

            // Выход из процедуры, если запрос пустой
            if (string.IsNullOrEmpty(SQLQuery))
            {
                MessageBox.Show("Вы не ввели запрос", "Отсутствует строка запроса", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверка, существует ли файл Movies.xlsx в той же папке, что и исполняемый файл
            movieFilePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Movies.xlsx");

            if (!System.IO.File.Exists(movieFilePath))
            {
                MessageBox.Show("Не удалось найти Movies.xlsx", "Файл не найден", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Создание и открытие соединения с файлом Movies.xlsx
                cn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={movieFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';");
                cn.Open();

                // Создание и заполнение DataTable с помощью SQL-запроса
                da = new OleDbDataAdapter(SQLQuery, cn);
                dt = new DataTable();
                da.Fill(dt);

                // Подсчет количества строк, возвращенных запросом
                rowCount = dt.Rows.Count;

                Console.WriteLine($"{rowCount} строк(и), {SQLQuery}");

                // Выход из процедуры, если результат запроса пуст
                if (rowCount == 0)
                {
                    MessageBox.Show("Запрос не вернул результатов", "Нет результатов", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Подсчет количества колонок, возвращенных запросом
                colCount = dt.Columns.Count;

                // Создание нового листа Excel
                ws = xlWorkbook.Sheets.Add();
                ws.Name = "QueryResults";

                // Форматирование заголовков таблицы
                Excel.Range headerRange = ws.Range["A1"].Resize[1, colCount];
                headerRange.Interior.Color = Excel.XlRgbColor.rgbCornflowerBlue;  // Заливка цветом
                headerRange.Font.Color = Excel.XlRgbColor.rgbWhite;  // Цвет шрифта
                headerRange.Font.Bold = true;  // Полужирный шрифт

                // Запись имен колонок в первую строку листа Excel
                for (int i = 0; i < colCount; i++)
                {
                    ws.Cells[1, i + 1] = dt.Columns[i].ColumnName;

                    // Применение специального формата для колонок с датами
                    if (dt.Columns[i].DataType == typeof(DateTime))
                    {
                        Excel.Range dateRange = ws.Range[ws.Cells[2, i + 1], ws.Cells[rowCount + 1, i + 1]];
                        dateRange.NumberFormat = "dd mmm yyyy";  // Формат даты
                    }
                }

                // Копирование данных из DataTable на лист Excel
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        ws.Cells[i + 2, j + 1] = dt.Rows[i][j];
                    }
                }

                // Автоподбор ширины колонок
                ws.Columns.AutoFit();
            }

            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                /*
                // Закрытие соединения и освобождение ресурсов
                if (cn != null) cn.Close();
                if (da != null) da.Dispose();
                if (dt != null) dt.Dispose();
                if (ws != null) xlWorkbook.Close(false);  // Закрытие без сохранения
                xlApp.Quit();
                */
            }

        }

        public void DeleteAllButMenuSheet_example()
        {
            /* 
            Эта процедура удаляет все листы в книге Excel, кроме листа с именем "MenuSheet".
            */

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "YourWorkbook.xlsx");
            Excel.Worksheet menuSheet = xlWorkbook.Sheets["MenuSheet"];

            // Отключение предупреждений перед удалением листов
            xlApp.DisplayAlerts = false;

            // Удаление всех листов, кроме "MenuSheet"
            foreach (Excel.Worksheet ws in xlWorkbook.Sheets)
            {
                if (ws.Name != menuSheet.Name)
                {
                    ws.Delete();
                }
            }

            // Включение предупреждений
            xlApp.DisplayAlerts = true;
            xlWorkbook.Save();  // Сохранение изменений в книге
            xlWorkbook.Close();  // Закрытие книги
            xlApp.Quit();  // Завершение работы Excel
        }
    }
}
