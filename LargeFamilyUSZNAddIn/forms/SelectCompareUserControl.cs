using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using LargeFamilyUSZNAddIn.classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LargeFamilyUSZNAddIn.forms
{
    public partial class SelectCompareUserControl : UserControl
    {
        WorksEcxel we = new WorksEcxel();

        public SelectCompareUserControl()
        {
            InitializeComponent();
            txtSamplingRangeTwo.Enabled = false;
            btRangeTwo.Enabled = false;
            btFIO.Enabled = false;  // Убрать, когда будет прописан функционал
            // установить значение по умолчанию
            cbTypeSelection.SelectedItem = cbTypeSelection.Items[0];

            //сделать невидемым
            tbСommentСell.Visible = false;
            butСommentСell.Visible=false;
            labSeparator.Visible = false;
            tbSeparator.Visible = false;
            butCompareForDifference.Visible = false;
            butOpenForm.Visible= false;

        }

        #region ВыборкаДиапазон

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender; // приводим отправителя к элементу типа CheckBox
            if (checkBox.Checked == true)
            {
                txtSamplingRangeTwo.Enabled = true;
                btRangeTwo.Enabled = true;
            }
            else
            {
                txtSamplingRangeTwo.Enabled = false;
                btRangeTwo.Enabled = false;
            }
        }

        private void btRangeOne_Click(object sender, EventArgs e)
        {
            string range = we.SelectRange();
            txtSamplingRangeOne.Text = range;
        }

        private void btRangeTwo_Click(object sender, EventArgs e)
        {
            string range = we.SelectRange();
            txtSamplingRangeTwo.Text = range;
        }
        #endregion

        #region ВыборкаФИО

        private void btFIO_Click(object sender, EventArgs e)
        {
            //TODO провисать функционал по выборке ФИО
        }

        private void btSelectRange_Click(object sender, EventArgs e)
        {
            string range = we.SelectRange();
            tbWhereFIO.Text = range;
        }


        #endregion

        #region ВыборкаЧисло

        //кнопка куда разместить выборку
        private void btSelectRangeNum_Click(object sender, EventArgs e)
        {
            string range = we.SelectRange();
            tbWhereNumber.Text = range;
            //TODO написать проверку если указан диапазон из нескольких ячеек
        }

        private void btNumberMask_Click(object sender, EventArgs e)
        {
            
            DataSampling ds = new DataSampling();
            List<string> cellTexts = new List<string>();
            List<string> recognizedNumbers = new List<string>();
            List<string> matchingMaskResult = new List<string>();
            List<string> nonMatchedResults = new List<string>();
            
            cellTexts = we.givenRangesList( txtSamplingRangeOne.Text);



            if (chbSeparator.Checked)
            {
                (matchingMaskResult, nonMatchedResults) = ds.FindNumbersByCustomPatternSeparator(cellTexts, tbMask.Text);
                if (tbChangeMask.Text == string.Empty)
                {
                    // разместить данные на листе Эксель
                    we.FillCells(tbWhereNumber.Text,matchingMaskResult,nonMatchedResults);
                }
                else
                {
                    List<string> transformedList = ds.TransformByMask(matchingMaskResult, tbChangeMask.Text);
                    we.FillCells(tbWhereNumber.Text, transformedList, nonMatchedResults);
                }
            }
            else
            {
                recognizedNumbers = ds.FindNumbersByCustomPattern(cellTexts, tbMask.Text);
                if (tbChangeMask.Text == string.Empty)
                {
                    // разместить данные на листе Эксель
                    we.FillCells(tbWhereNumber.Text, recognizedNumbers);
                }
                else
                {
                    List<string> transformedList = ds.TransformByMask(recognizedNumbers, tbChangeMask.Text);
                    we.FillCells(tbWhereNumber.Text, transformedList);
                }

            }  
           
        }

        #endregion
        #region СравнениеЧисло
        private void btCompareForMatch_Click(object sender, EventArgs e)
        {
            
            DataSampling ds = new DataSampling();
            Dictionary<string, string> cellTextsA = new Dictionary<string, string>();
            Dictionary<string, string> cellTextsB = new Dictionary<string, string>();
            //выборка данных для сравнения
            //Dictionary<string, XLWorkbook> workbook = we.LoadAllOpenWorkbooksDesk();
            cellTextsA = we.givenRangesDictionary(txtSamplingRangeOne.Text);
            cellTextsB = we.givenRangesDictionary(txtSamplingRangeTwo.Text);
            //обработка выбранных данных
            Dictionary<string, string> recognizedNumbersA = ds.FindNumbersByCustomPatternDictionary(cellTextsA, tbMaskSelection.Text);
            Dictionary<string, string> recognizedNumbersB = ds.FindNumbersByCustomPatternDictionary(cellTextsB, txMaskSearch.Text);
            List<string> resultСomparison = new List<string>();
            //Проверяем нужно менять даные перед сравнением
            if (tbtMaskСomparison.Text != string.Empty)
            {
                Dictionary<string, string> transformedDictionary = ds.TransformByMaskDictionary(recognizedNumbersB, tbtMaskСomparison.Text);
                //Сравнение данных
                resultСomparison = ds.CompareDictionaries(recognizedNumbersA, transformedDictionary);
            }
            else 
            {
                resultСomparison = ds.CompareDictionaries(recognizedNumbersA, recognizedNumbersB);
            }
            //окраска ячеек 
            we.ColorCellsInPink(resultСomparison);
            
        }
        #endregion
    }
}
