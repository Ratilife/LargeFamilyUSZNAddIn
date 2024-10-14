using LargeFamilyUSZNAddIn.forms;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LargeFamilyUSZNAddIn
{
    public partial class RibbonSC
    {
        private void RibbonSC_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void butSC_Click(object sender, RibbonControlEventArgs e)
        {
            SelectCompareUserControl form = new SelectCompareUserControl();
            Microsoft.Office.Tools.CustomTaskPane customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(form, "Выбрать/Сравнить");
            AdjustCustomTaskPaneSize(customTaskPane);
            customTaskPane.Visible = true;

        }

        private void AdjustCustomTaskPaneSize(Microsoft.Office.Tools.CustomTaskPane customTaskPane)
        {
            // Получаем размер экрана
            int screenWidth = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;

            // Устанавливаем ширину формы пропорционально размеру экрана
            customTaskPane.Width = screenWidth / 4; // Например, треть экрана

            // Опционально: можно установить минимально и максимально допустимые размеры панели задач
            customTaskPane.Width = Math.Max(200, Math.Min(screenWidth / 3, 400)); // Ограничение по ширине

        }
    }
}
