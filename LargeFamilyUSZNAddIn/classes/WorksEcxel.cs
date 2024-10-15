using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace LargeFamilyUSZNAddIn.classes
{
    internal class WorksEcxel
    {
        private IXLWorkbook workbook;
        public WorksEcxel() 
        {
            // Загружаем активную книгу Excel
            this.LoadActiveWorkbook();
        }


        #region ВыборкаОбщая

        // Метод для загрузки активной книги Excel в ClosedXML
        /**
        * Метод LoadActiveWorkbook загружает активную книгу Excel в формате .xlsx для работы с ней через библиотеку ClosedXML.
        * 
        * Основные шаги метода:
        * 1. Попытка получить текущий экземпляр приложения Excel через Globals.ThisAddIn.Application.
        *    - Если Excel не открыт, выбрасывается COMException, и выводится сообщение об ошибке.
        * 2. Проверка наличия активной книги:
        *    - Если активная книга отсутствует (нет открытых книг в Excel), метод завершает выполнение с сообщением "Нет активной книги".
        * 3. Временное сохранение активной книги как .xlsx файл:
        *    - Для этого создается временный файл в директории Temp (системная временная директория), используя метод `SaveCopyAs`.
        * 4. Загрузка временного файла в объект XLWorkbook библиотеки ClosedXML:
        *    - Это позволяет работать с книгой Excel через функционал библиотеки ClosedXML.
        * 5. После успешной загрузки книга сохраняется в переменной `this.workbook`, а временный файл удаляется, чтобы не оставлять его на диске.
        *    - Удаление происходит через метод `System.IO.File.Delete`.
        * 
        * Метод полезен в случаях, когда нужно работать с активной книгой Excel с использованием возможностей библиотеки ClosedXML,
        * при этом Excel-книга временно сохраняется как файл .xlsx, чтобы можно было ее загрузить в ClosedXML для последующей обработки.
        */

        private void LoadActiveWorkbook()
        {
            Excel.Application excelApp;

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

            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            if (activeWorkbook == null)
            {
                MessageBox.Show("Нет активной книги");
                return;
            }

            // Временно сохраняем активную книгу как .xlsx файл и загружаем в ClosedXML
            string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "tempExcelFile.xlsx");
            activeWorkbook.SaveCopyAs(tempFilePath);

            // Загружаем файл в ClosedXML
            this.workbook = new XLWorkbook(tempFilePath);

            // Удаляем временный файл после загрузки
            System.IO.File.Delete(tempFilePath);
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

            // Имя листа
            string sheetName = fileNameAndSheet[1].Split('\'')[1].Trim();

            // Извлекаем диапазон ячеек
            var cellRange = fileNameAndSheet[1].Split('\'')[2].Trim(); // Диапазон ячеек

            foreach (var ws in workbook.Worksheets)
            {
                Console.WriteLine($"Лист: {ws.Name}");

            }
                // Находим лист по имени
                var worksheet = workbook.Worksheet(sheetName);
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
                        // Проверяем, что значение ячейки не пустое
                        //if (!string.IsNullOrEmpty(cellValue))
                        //{
                            //cellTexts.Add(null);
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении диапазона: {ex.Message}");
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
       
        #endregion

    }
}
