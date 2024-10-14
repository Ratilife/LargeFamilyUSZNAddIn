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
        public SelectCompareUserControl()
        {
            InitializeComponent();
            txtSamplingRangeTwo.Enabled = false;
            btRangeTwo.Enabled = false;
        }

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
    }
}
