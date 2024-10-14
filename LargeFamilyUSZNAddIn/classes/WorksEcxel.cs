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
        #region ВыборкаОбщая

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

            Excel.Worksheet activeSheet = workbook.ActiveSheet;
            if (activeSheet == null)
            {
                MessageBox.Show("Нет активного листа");
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

            string selectedAddress = selectedRange.get_Address();

            // Определяем имя листа, на котором выделен диапазон
            Excel.Worksheet selectedSheet = selectedRange.Worksheet;
            if (selectedSheet != activeSheet)
            {
                selectedAddress = selectedSheet.Name + "!" + selectedAddress;
            }

            return selectedAddress;
        }

        #endregion

        public List<string> givenRangesList(IXLWorkbook workbook, string rangeAddress)
        {
            List<string> cellTexts = new List<string>();

            // Извлекаем имя листа из диапазона
            var rangeParts = rangeAddress.Split('!');
            var sheetName = rangeParts[0].Replace("[", "").Replace("]", ""); // Извлекаем имя листа
            var cellRange = rangeParts[1]; // Извлекаем диапазон ячеек

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


    }
}
