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

                        // Проверяем, что значение ячейки не пустое
                        if (!string.IsNullOrEmpty(cellValue))
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

            return cellTexts;
        }

        //проверить
        public List<string> givenRangesList2(string rangeAddress)
        {
            List<string> cellTexts = new List<string>();

            // Извлекаем имя файла и листа из диапазона
            var rangeParts = rangeAddress.Split('!');

            // Имя файла в скобках (мы их не используем)
            string fullFileName = rangeParts[0].Replace("[", "").Replace("]", "").Trim();

            // Имя листа
            string sheetName = fullFileName.Substring(fullFileName.IndexOf(']') + 1).Trim();

            // Извлекаем диапазон ячеек
            var cellRange = rangeParts[1]; // Диапазон ячеек

            // Находим лист по имени
            var worksheet = workbook.Worksheet(sheetName);

            // Получаем диапазон ячеек
            var range = worksheet.Range(cellRange);

            // Перебираем строки и ячейки в указанном диапазоне
            foreach (var row in range.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    cellTexts.Add(cell.GetString());
                }
            }

            return cellTexts;
        }

        //TODO переписать метод  
        public void FillCells(string startCellAddress, List<string> dataList)
        {
            // Извлекаем имя файла и листа из диапазона
            var fileNameAndSheet = startCellAddress.Split(new[] { ']' }, 2);
            if (fileNameAndSheet.Length < 2)
            {
                MessageBox.Show("Некорректный адрес диапазона. Убедитесь, что включена часть с диапазоном.");
            }

            // Имя файла
            string fullFileName = fileNameAndSheet[0].Replace("[", "").Trim();

            // Имя листа
            string sheetName = fileNameAndSheet[1].Split('\'')[1].Trim();

            // Извлекаем диапазон ячеек
            var cellRange = fileNameAndSheet[1].Split('\'')[2].Trim(); // Диапазон ячеек
            
            var worksheet = workbook.Worksheet(sheetName); // Или измените на нужный лист

            // Находим ячейку, с которой начинаем
            var startCell = worksheet.Cell(cellRange);

            int startRow = startCell.Address.RowNumber; // Получаем номер строки
            int startColumn = startCell.Address.ColumnNumber; // Получаем номер столбца

            // Заполняем ячейки данными из списка
            for (int i = 0; i < dataList.Count; i++)
            {
                var cell = worksheet.Cell(startRow + i, startColumn); // Находим нужную ячейку
                cell.Value = dataList[i]; // Заполняем ячейку значением из списка
            }
        }
    }
}
