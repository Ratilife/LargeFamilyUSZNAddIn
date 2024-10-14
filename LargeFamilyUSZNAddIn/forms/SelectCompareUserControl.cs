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
            var recognizedNumbers = ds.FindNumbersByCustomPattern(cellTexts, tbMask.Text);
        }



        #endregion

        
    }
}
