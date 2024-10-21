using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using CsvHelper;
using System.Text.RegularExpressions;



namespace LargeFamilyUSZNAddIn.classes
{
    internal class WorksEcxel
    {
        //private IXLWorkbook workbook;
        private // Создаем список для хранения загруженных книг
        Dictionary <string, XLWorkbook> workbooksDictionary = new Dictionary<string, XLWorkbook>();

        public WorksEcxel() 
        {
            
        }


        #region ВыборкаОбщая

        private void LoadAllOpenWorkbooks()
        {
            Excel.Application excelApp;
            Excel.Workbooks workbooks;

            try
            {
                // Получаем текущий экземпляр приложения Excel
                excelApp = Globals.ThisAddIn.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("Excel не открыт");
                return;
            }

            // Получаем все открытые книги
            workbooks = excelApp.Workbooks;

            if (workbooks.Count == 0)
            {
                MessageBox.Show("Нет открытых книг Excel.");
                return;
            }

            foreach (Excel.Workbook wb in workbooks)
            {
                try
                {
                    // Создаем временный файл
                    string tempFilePath = Path.GetTempFileName();
                    // Имя файла
                    string fileName = wb.Name;
                    // Сохраняем книгу во временный файл
                    wb.SaveCopyAs(tempFilePath);

                    // Открываем временный файл в ClosedXML
                    using (var stream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                         var workbook = new XLWorkbook(stream);

                        // Добавляем загруженный workbook в список
                        workbooksDictionary.Add(fileName, workbook);
                    }

                    // Удаляем временный файл
                    File.Delete(tempFilePath);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Ошибка при обработке книги {wb.Name}: {ex.Message}");
                }
            }
        }

       

        public string SelectRange()
        {
            Excel.Application excelApp;
            object inputBoxResult;

            try
            {
                // Получение текущего экземпляра приложения Excel
                excelApp = Globals.ThisAddIn.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("Excel не открыт");
                return null;
            }

            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                MessageBox.Show("Нет активной книги");
                return null;
            }

            inputBoxResult = excelApp.InputBox("Выберите диапазон", Type: 8);
            if (inputBoxResult is bool && (bool)inputBoxResult == false)
            {
                MessageBox.Show("Диапазон не выбран");
                return null;
            }

            Excel.Range selectedRange = (Excel.Range)inputBoxResult;

            // Проверка, что диапазон не превышает один столбец
            if (selectedRange.Columns.Count > 1)
            {
                MessageBox.Show("Нельзя выбирать диапазон больше одного столбца");
                return null;
            }

            // Получаем полный адрес
            string selectedAddress = selectedRange.get_Address();

            // Получаем имя книги и листа
            Excel.Worksheet selectedSheet = selectedRange.Worksheet;
            string workbookName = selectedSheet.Parent.Name;

            // Формируем полный адрес с именем файла
            selectedAddress = $"[{workbookName}]'{selectedSheet.Name}'{selectedAddress}";

            return selectedAddress;
        }


        #endregion
       
        #region ВыборкаЧисел
        public List<string> givenRangesList(string rangeAddress)
        {
            LoadAllOpenWorkbooks();
            List<string> cellTexts = new List<string>();
            

            // Проверяем, содержит ли строка правильный формат с диапазоном
            if (!rangeAddress.Contains("'"))
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что адрес формирован верно.");
                return cellTexts; // Возвращаем пустой список
            }

            // Извлекаем имя файла и листа из диапазона
            var fileNameAndSheet = rangeAddress.Split(new[] { ']' }, 2);

            if (fileNameAndSheet.Length < 2)
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что включена часть с диапазоном.");
                return cellTexts; // Возвращаем пустой список
            }

            // Имя файла
            string fullFileName = fileNameAndSheet[0].Replace("[", "").Trim();

            // Проверка расширения файла
            if (fullFileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                // Для файлов .csv пропускаем код, который работает с листами Excel
                MessageBox.Show("Обработка файлов .csv на данный момент не поддерживается.");
                return cellTexts;
            }

            // Имя листа
            string sheetName = fileNameAndSheet[1].Split('\'')[1].Trim();

            // Извлекаем диапазон ячеек
            var cellRange = fileNameAndSheet[1].Split('\'')[2].Trim(); // Диапазон ячеек

            // Определяем worksheet как null вначале
            //IXLWorksheet worksheet = null;
            // Получаем все открытые книги
            
            if (fullFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                // Перебираем все книги в workbooksList
                foreach (KeyValuePair < string, XLWorkbook > workbook in workbooksDictionary)
                {

                    // Проверяем, совпадает ли имя файла с полным именем
                    //if (workbook.FullPath == fullFileName) // Убедитесь, что FullPath доступен
                    //{
                        // Проходим по всем листам в текущей книге
                        foreach (IXLWorksheet worksheet in workbook.Value.Worksheets)
                        {
                            // Проверяем, совпадает ли имя листа с заданным именем
                            if (worksheet.Name == sheetName)
                            {
                                if (worksheet == null)
                                {
                                    MessageBox.Show($"Лист с именем '{sheetName}' не найден.");
                                    return cellTexts; // Возвращаем пустой список
                                }



                                try
                                {
                                    // Получаем диапазон ячеек
                                    var range = worksheet.Range(cellRange);

                                    // Перебираем строки и ячейки в указанном диапазоне
                                    foreach (var row in range.Rows())
                                    {
                                        foreach (var cell in row.Cells())
                                        {
                                            string cellValue = cell.GetString();

                                            var texts = new List<string>
                                {
                                    cellValue
                                };

                                            if (texts.Count > 1)
                                            {
                                                foreach (var number in texts)
                                                {
                                                    cellTexts.Add(number);
                                                }
                                            }
                                            else
                                            {
                                                cellTexts.Add(cellValue);
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Ошибка при получении диапазона: {ex.Message}");
                                }
                            }
                        }
                    //}
                }
            } 
            return cellTexts;
        }



        /**
        * Метод FillCells предназначен для заполнения ячеек на листе Excel данными из списка.
        * 
        * Описание работы метода:
        * 1. Метод принимает два параметра:
        *    - startCellAddress: строка, содержащая адрес начальной ячейки, а также имя файла и листа, например: "[Файл.xlsx]'Лист1'$A$1".
        *    - dataList: список строк, содержащий данные для записи в ячейки.
        * 
        * 2. В начале метод пытается получить активное приложение Excel через объект Globals.ThisAddIn.Application.
        *    В случае, если Excel не запущен, возникает COMException, и пользователю выводится сообщение об ошибке.
        * 
        * 3. Далее происходит разбор адреса startCellAddress:
        *    - Извлекается имя файла и листа из строки.
        *    - Если имя файла не указано или указано как "Книга1" (несохраненная книга), метод работает с активной книгой. Если активной книги нет, выводится сообщение об ошибке.
        *    - Если указано имя файла, метод ищет его среди открытых книг. Если нужная книга не найдена, выводится сообщение об ошибке.
        * 
        * 4. После этого метод ищет указанный лист в выбранной книге.
        *    Если лист с указанным именем не найден, выводится сообщение об ошибке.
        * 
        * 5. Метод находит начальную ячейку на листе, начиная с которой будут записываться данные.
        *    Из адреса ячейки извлекаются номер строки и номер столбца.
        * 
        * 6. Далее происходит запись данных из списка dataList в соответствующие ячейки.
        *    Для каждой строки списка dataList метод последовательно находит нужную ячейку и записывает в нее значение.
        * 
        * 7. Сохранение файла в методе не выполняется, так как это должен решить пользователь.
        *    Метод предназначен только для записи данных в открытые книги.
        */
        public void FillCells(string startCellAddress, List<string> dataList)
        {
            Excel.Application excelApp;
            Excel.Workbook targetWorkbook = null;

            try
            {
                // Получаем текущий экземпляр приложения Excel
                excelApp = Globals.ThisAddIn.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("Excel не открыт");
                return;
            }

            // Извлекаем имя файла и листа из диапазона
            var fileNameAndSheet = startCellAddress.Split(new[] { ']' }, 2);
            if (fileNameAndSheet.Length < 2)
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что включена часть с диапазоном.");
                return;
            }

            // Имя файла
            string fullFileName = fileNameAndSheet[0].Replace("[", "").Trim();
            // Имя листа
            string sheetName = fileNameAndSheet[1].Split('\'')[1].Trim();
            // Адрес ячейки
            string cellRange = fileNameAndSheet[1].Split('\'')[2].Trim();

            // Проверяем, указан ли файл
            if (string.IsNullOrEmpty(fullFileName) || fullFileName.Equals("Книга1", StringComparison.OrdinalIgnoreCase))
            {
                // Если файл не указан или указан как "Книга1", считаем его активной книгой
                targetWorkbook = excelApp.ActiveWorkbook;
                if (targetWorkbook == null)
                {
                    MessageBox.Show("Нет активной книги.");
                    return;
                }
            }
            else
            {
                // Ищем среди открытых книг указанную
                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    if (wb.Name.Equals(fullFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetWorkbook = wb;
                        break;
                    }
                }

                if (targetWorkbook == null)
                {
                    MessageBox.Show($"Файл '{fullFileName}' не найден среди открытых книг.");
                    return;
                }
            }

            // Получаем лист по имени
            Excel.Worksheet worksheet = null;
            foreach (Excel.Worksheet sheet in targetWorkbook.Sheets)
            {
                if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
            {
                MessageBox.Show($"Лист с именем '{sheetName}' не найден.");
                return;
            }

            // Находим начальную ячейку
            Excel.Range startCell = worksheet.Range[cellRange];
            int startRow = startCell.Row;
            int startColumn = startCell.Column;


            // Заполняем ячейки данными из списка
            for (int i = 0; i < dataList.Count; i++)
            {
                Excel.Range cell = worksheet.Cells[startRow + i, startColumn];
                cell.Value = dataList[i];
            }

            // Сохранять файл не нужно, так как пользователь сам решает, сохранять или нет
        }

        /*
        * Заполняет ячейки в указанной книге Excel данными из двух списков.
        *
        * Параметры:
        * - startCellAddress: Адрес начальной ячейки в формате "[ИмяФайла]ИмяЛиста'НачальнаяЯчейка'".
        * - dataList1: Список строк, которые будут записаны в первый столбец, начиная с указанной начальной ячейки.
        * - dataList2: Список строк, которые будут записаны в соседний столбец справа от первого, начиная с той же строки.
        *
        * Описание:
        * Метод выполняет следующие шаги:
        * 1. Получает текущий экземпляр приложения Excel.
        * 2. Извлекает имя файла и имя листа из переданного адреса начальной ячейки.
        * 3. Определяет, какой файл (книга) использовать — активную книгу или указанную в адресе.
        * 4. Находит указанный лист в целевой книге.
        * 5. Находит начальную ячейку и определяет, в какие столбцы будут записываться данные.
        * 6. Заполняет ячейки первого списка данными, начиная с указанной начальной ячейки.
        * 7. Заполняет соседний столбец данными из второго списка, начиная с той же строки.
        *
        * Метод не сохраняет изменения в файле, оставляя это решение за пользователем.
        * Если возникнут ошибки, такие как отсутствие активной книги или указанного листа,
        * будет показано соответствующее сообщение.
        */

        public void FillCells(string startCellAddress, List<string> dataList1, List<string> dataList2)
        {
            Excel.Application excelApp;
            Excel.Workbook targetWorkbook = null;

            try
            {
                // Получаем текущий экземпляр приложения Excel
                excelApp = Globals.ThisAddIn.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("Excel не открыт");
                return;
            }

            // Извлекаем имя файла и листа из диапазона
            var fileNameAndSheet = startCellAddress.Split(new[] { ']' }, 2);
            if (fileNameAndSheet.Length < 2)
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что включена часть с диапазоном.");
                return;
            }

            // Имя файла
            string fullFileName = fileNameAndSheet[0].Replace("[", "").Trim();
            // Имя листа
            string sheetName = fileNameAndSheet[1].Split('\'')[1].Trim();
            // Адрес ячейки
            string cellRange = fileNameAndSheet[1].Split('\'')[2].Trim();

            // Проверяем, указан ли файл
            if (string.IsNullOrEmpty(fullFileName) || fullFileName.Equals("Книга1", StringComparison.OrdinalIgnoreCase))
            {
                // Если файл не указан или указан как "Книга1", считаем его активной книгой
                targetWorkbook = excelApp.ActiveWorkbook;
                if (targetWorkbook == null)
                {
                    MessageBox.Show("Нет активной книги.");
                    return;
                }
            }
            else
            {
                // Ищем среди открытых книг указанную
                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    if (wb.Name.Equals(fullFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetWorkbook = wb;
                        break;
                    }
                }

                if (targetWorkbook == null)
                {
                    MessageBox.Show($"Файл '{fullFileName}' не найден среди открытых книг.");
                    return;
                }
            }

            // Получаем лист по имени
            Excel.Worksheet worksheet = null;
            foreach (Excel.Worksheet sheet in targetWorkbook.Sheets)
            {
                if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
            {
                MessageBox.Show($"Лист с именем '{sheetName}' не найден.");
                return;
            }

            // Находим начальную ячейку
            Excel.Range startCell = worksheet.Range[cellRange];
            int startRow = startCell.Row;
            int startColumn = startCell.Column;

            // Убедимся, что длина списков совпадает
            int maxRows = Math.Max(dataList1.Count, dataList2.Count);

            // Заполняем данные первого списка в начальный столбец
            for (int i = 0; i < dataList1.Count; i++)
            {
                Excel.Range cell = worksheet.Cells[startRow + i, startColumn];
                cell.Value = dataList1[i];
            }

            // Заполняем данные второго списка в соседний столбец
            int adjacentColumn = startColumn + 1;  // Соседний столбец справа
            for (int i = 0; i < dataList2.Count; i++)
            {
                Excel.Range cell = worksheet.Cells[startRow + i, adjacentColumn];
                cell.Value = dataList2[i];
            }

            // Сохранять файл не нужно, так как пользователь сам решает, сохранять или нет
        }

        #endregion

        #region СравнениеЧисел
        public Dictionary<string, string> givenRangesDictionary(string rangeAddress)
        {
            LoadAllOpenWorkbooks();
            // Создаем словарь для хранения результата
            Dictionary<string, string> cellDataDict = new Dictionary<string, string>();

            // Проверяем, содержит ли строка правильный формат с диапазоном
            if (!rangeAddress.Contains("'"))
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что адрес формирован верно.");
                return cellDataDict; // Возвращаем пустой словарь
            }

            // Извлекаем имя файла и листа из диапазона
            var fileNameAndSheet = rangeAddress.Split(new[] { ']' }, 2);

            if (fileNameAndSheet.Length < 2)
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что включена часть с диапазоном.");
                return cellDataDict; // Возвращаем пустой словарь
            }

            // Имя файла
            string fullFileName = fileNameAndSheet[0].Replace("[", "").Trim();

            // Проверка расширения файла
            if (fullFileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                // Для файлов .csv пропускаем код, который работает с листами Excel
                MessageBox.Show("Обработка файлов .csv на данный момент не поддерживается.");
                return cellDataDict;
            }

            // Имя листа
            string sheetName = fileNameAndSheet[1].Split('\'')[1].Trim();

            // Извлекаем диапазон ячеек
            var cellRange = fileNameAndSheet[1].Split('\'')[2].Trim(); // Диапазон ячеек

            //получаем все открытые книги 
            if (fullFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                foreach (KeyValuePair<string, XLWorkbook> workbook in workbooksDictionary)
                {
                    // Проходим по всем листам в текущей книге
                    foreach (IXLWorksheet worksheet in workbook.Value.Worksheets)
                    {
                        if (worksheet.Name == sheetName)
                        {
                            if (worksheet == null)
                            {
                                MessageBox.Show($"Лист с именем '{sheetName}' не найден.");
                                return cellDataDict; // Возвращаем пустой список
                            }
                            try
                            {
                                // Получаем диапазон ячеек
                                var range = worksheet.Range(cellRange);

                                // Переменная для отслеживания порядкового номера
                                int index = 1;

                                // Перебираем строки и ячейки в указанном диапазоне
                                foreach (var row in range.Rows())
                                {
                                    foreach (var cell in row.Cells())
                                    {
                                        string cellValue = cell.GetString();

                                        // Проверка на наличие значений в ячейке
                                        if (!string.IsNullOrEmpty(cellValue))
                                        {
                                            // Формируем ключ в формате [имяФайла]'ИмяЛиста'A1(1)
                                            string key = $"[{fullFileName}]'{sheetName}'{cell.Address}({index})";
                                            cellDataDict.Add(key, cellValue);
                                            index++; // Увеличиваем индекс для каждого значения
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Ошибка при получении диапазона: {ex.Message}");
                            }

                        }
                    }
                }
                    /*// Перебираем все книги в workbooksDictionary
                    foreach (var ws in workbook.Worksheets)
                    {
                        Console.WriteLine($"Лист: {ws.Name}");
                    }

                    // Находим лист по имени
                    var worksheet = workbook.Worksheet(sheetName);
                    if (worksheet == null)
                    {
                        MessageBox.Show($"Лист с именем '{sheetName}' не найден.");
                        return cellDataDict; // Возвращаем пустой словарь
                    }

                    try
                    {
                        // Получаем диапазон ячеек
                        var range = worksheet.Range(cellRange);

                        // Переменная для отслеживания порядкового номера
                        int index = 1;

                        // Перебираем строки и ячейки в указанном диапазоне
                        foreach (var row in range.Rows())
                        {
                            foreach (var cell in row.Cells())
                            {
                                string cellValue = cell.GetString();

                                // Проверка на наличие значений в ячейке
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    // Формируем ключ в формате [имяФайла]'ИмяЛиста'A1(1)
                                    string key = $"[{fullFileName}]'{sheetName}'{cell.Address}(index)";
                                    cellDataDict.Add(key, cellValue);
                                    index++; // Увеличиваем индекс для каждого значения
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при получении диапазона: {ex.Message}");
                    }*/
            }
            return cellDataDict; // Возвращаем заполненный словарь
        }

        /* 
            Метод ColorCellsInPink выполняет окрашивание указанных ячеек в розовый цвет в открытых книгах Excel на основании переданного списка строк с информацией о ячейках.

            1. Получает текущий экземпляр приложения Excel через `Globals.ThisAddIn.Application` и доступ ко всем открытым книгам Excel.

            2. Для каждой строки в списке `cellInfoList` с помощью регулярного выражения извлекает:
               - Имя файла (книги Excel),
               - Имя листа,
               - Адрес ячейки.

            3. В цикле ищет книгу Excel по имени файла среди всех открытых книг. Если книга найдена, переходит к поиску указанного листа внутри этой книги.

            4. Если указанный лист найден, метод извлекает диапазон ячеек на основе адреса ячейки и устанавливает для них розовый цвет с помощью свойства `Interior.Color`.

            5. Если указанный файл или лист не найдены, выводятся сообщения об ошибке в консоль, чтобы уведомить пользователя о проблеме (например, если имя файла или листа указаны неверно).

            6. Если формат строки не соответствует ожидаемому шаблону, метод выводит сообщение об ошибке, информируя пользователя о некорректных данных.
       */

        public void ColorCellsInPink(List<string> cellInfoList)
        {
            Excel.Application excelApp;
            Excel.Workbooks workbooks;
            
            // Получаем текущий экземпляр приложения Excel
            excelApp = Globals.ThisAddIn.Application;
            // Получаем все открытые книги
            workbooks = excelApp.Workbooks;
            // Регулярное выражение для извлечения данных: имя файла, имя листа, адрес ячейки
            string pattern = @"\[(.*?)\]'(.*?)'([A-Z]+\d+)";
            Regex regex = new Regex(pattern);

            foreach (string cellInfo in cellInfoList)
            {
                Match match = regex.Match(cellInfo);

                if (match.Success)
                {
                    // Извлекаем информацию из строки
                    string fileName = match.Groups[1].Value;  // Имя файла
                    string sheetName = match.Groups[2].Value; // Имя листа
                    string cellAddress = match.Groups[3].Value; // Адрес ячейки

                    // Ищем открытую книгу Excel по имени файла
                    Excel.Workbook targetWorkbook = null;
                    foreach (Excel.Workbook workbook in workbooks)
                    {
                        // Проверяем имя книги
                        if (workbook.Name == fileName)
                        {
                            targetWorkbook = workbook;
                            break;
                        }
                    }

                    if (targetWorkbook != null)
                    {
                        // Ищем лист по имени
                        Excel.Worksheet worksheet = null;
                        foreach (Excel.Worksheet ws in targetWorkbook.Sheets)
                        {
                            if (ws.Name == sheetName)
                            {
                                worksheet = ws;
                                break;
                            }
                        }

                        if (worksheet != null)
                        {
                            // Окрашиваем ячейку в розовый цвет
                            Excel.Range cell = worksheet.get_Range(cellAddress);
                            cell.Interior.Color = Excel.XlRgbColor.rgbPink;
                        }
                        else
                        {
                            Console.WriteLine($"Ошибка: лист '{sheetName}' не найден в книге '{fileName}'.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Ошибка: книга '{fileName}' не найдена среди открытых.");
                    }
                }
                else
                {
                    Console.WriteLine($"Ошибка: строка '{cellInfo}' не соответствует формату.");
                }
            }
        }

        
        #endregion
    }
}
