namespace LargeFamilyUSZNAddIn.forms
{
    partial class SelectCompareUserControl
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btRangeTwo = new System.Windows.Forms.Button();
            this.txtSamplingRangeTwo = new System.Windows.Forms.TextBox();
            this.btRangeOne = new System.Windows.Forms.Button();
            this.txtSamplingRangeOne = new System.Windows.Forms.TextBox();
            this.tbSelectionCompare = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.tbChangeMask = new System.Windows.Forms.TextBox();
            this.lbTextChange = new System.Windows.Forms.Label();
            this.btNumberMask = new System.Windows.Forms.Button();
            this.btSelectRangeNum = new System.Windows.Forms.Button();
            this.tbWhereNumber = new System.Windows.Forms.TextBox();
            this.lblWhereNumber = new System.Windows.Forms.Label();
            this.tbMask = new System.Windows.Forms.TextBox();
            this.labelTextMask = new System.Windows.Forms.Label();
            this.labelHeadingMaskSearch = new System.Windows.Forms.Label();
            this.btSelectRange = new System.Windows.Forms.Button();
            this.tbWhereFIO = new System.Windows.Forms.TextBox();
            this.lblWhereFIO = new System.Windows.Forms.Label();
            this.btFIO = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.butСommentСell = new System.Windows.Forms.Button();
            this.tbСommentСell = new System.Windows.Forms.TextBox();
            this.labСommentСell = new System.Windows.Forms.Label();
            this.tbtMaskСomparison = new System.Windows.Forms.TextBox();
            this.txMaskSearch = new System.Windows.Forms.TextBox();
            this.tbMaskSelection = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.labTextMaskСomparison = new System.Windows.Forms.Label();
            this.labTextMaskSearch = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tbSeparator = new System.Windows.Forms.TextBox();
            this.labSeparator = new System.Windows.Forms.Label();
            this.cbTypeSelection = new System.Windows.Forms.ComboBox();
            this.labTypeSelection = new System.Windows.Forms.Label();
            this.tbSelectionCompare.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(18, 70);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(104, 17);
            this.checkBox1.TabIndex = 18;
            this.checkBox1.Text = "Для сравнения";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "Выбрать диапазон";
            // 
            // btRangeTwo
            // 
            this.btRangeTwo.Location = new System.Drawing.Point(344, 70);
            this.btRangeTwo.Name = "btRangeTwo";
            this.btRangeTwo.Size = new System.Drawing.Size(28, 23);
            this.btRangeTwo.TabIndex = 16;
            this.btRangeTwo.Text = "...";
            this.btRangeTwo.UseVisualStyleBackColor = true;
            this.btRangeTwo.Click += new System.EventHandler(this.btRangeTwo_Click);
            // 
            // txtSamplingRangeTwo
            // 
            this.txtSamplingRangeTwo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSamplingRangeTwo.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.txtSamplingRangeTwo.Location = new System.Drawing.Point(144, 70);
            this.txtSamplingRangeTwo.Name = "txtSamplingRangeTwo";
            this.txtSamplingRangeTwo.Size = new System.Drawing.Size(194, 20);
            this.txtSamplingRangeTwo.TabIndex = 15;
            // 
            // btRangeOne
            // 
            this.btRangeOne.Location = new System.Drawing.Point(344, 28);
            this.btRangeOne.Name = "btRangeOne";
            this.btRangeOne.Size = new System.Drawing.Size(28, 23);
            this.btRangeOne.TabIndex = 14;
            this.btRangeOne.Text = "...";
            this.btRangeOne.UseVisualStyleBackColor = true;
            this.btRangeOne.Click += new System.EventHandler(this.btRangeOne_Click);
            // 
            // txtSamplingRangeOne
            // 
            this.txtSamplingRangeOne.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSamplingRangeOne.Location = new System.Drawing.Point(144, 30);
            this.txtSamplingRangeOne.Name = "txtSamplingRangeOne";
            this.txtSamplingRangeOne.Size = new System.Drawing.Size(194, 20);
            this.txtSamplingRangeOne.TabIndex = 13;
            // 
            // tbSelectionCompare
            // 
            this.tbSelectionCompare.Controls.Add(this.tabPage1);
            this.tbSelectionCompare.Controls.Add(this.tabPage2);
            this.tbSelectionCompare.Location = new System.Drawing.Point(18, 104);
            this.tbSelectionCompare.Name = "tbSelectionCompare";
            this.tbSelectionCompare.SelectedIndex = 0;
            this.tbSelectionCompare.ShowToolTips = true;
            this.tbSelectionCompare.Size = new System.Drawing.Size(354, 445);
            this.tbSelectionCompare.TabIndex = 19;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.tbChangeMask);
            this.tabPage1.Controls.Add(this.lbTextChange);
            this.tabPage1.Controls.Add(this.btNumberMask);
            this.tabPage1.Controls.Add(this.btSelectRangeNum);
            this.tabPage1.Controls.Add(this.tbWhereNumber);
            this.tabPage1.Controls.Add(this.lblWhereNumber);
            this.tabPage1.Controls.Add(this.tbMask);
            this.tabPage1.Controls.Add(this.labelTextMask);
            this.tabPage1.Controls.Add(this.labelHeadingMaskSearch);
            this.tabPage1.Controls.Add(this.btSelectRange);
            this.tabPage1.Controls.Add(this.tbWhereFIO);
            this.tabPage1.Controls.Add(this.lblWhereFIO);
            this.tabPage1.Controls.Add(this.btFIO);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(346, 419);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Выборка";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(216, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 13);
            this.label2.TabIndex = 27;
            this.label2.Text = "# - это любое число";
            // 
            // tbChangeMask
            // 
            this.tbChangeMask.Location = new System.Drawing.Point(78, 135);
            this.tbChangeMask.Name = "tbChangeMask";
            this.tbChangeMask.Size = new System.Drawing.Size(210, 20);
            this.tbChangeMask.TabIndex = 26;
            // 
            // lbTextChange
            // 
            this.lbTextChange.AutoSize = true;
            this.lbTextChange.Location = new System.Drawing.Point(15, 139);
            this.lbTextChange.Name = "lbTextChange";
            this.lbTextChange.Size = new System.Drawing.Size(61, 13);
            this.lbTextChange.TabIndex = 25;
            this.lbTextChange.Text = "Изменить ";
            // 
            // btNumberMask
            // 
            this.btNumberMask.Location = new System.Drawing.Point(20, 209);
            this.btNumberMask.Name = "btNumberMask";
            this.btNumberMask.Size = new System.Drawing.Size(140, 23);
            this.btNumberMask.TabIndex = 24;
            this.btNumberMask.Text = "Выборка по маске";
            this.btNumberMask.UseVisualStyleBackColor = true;
            this.btNumberMask.Click += new System.EventHandler(this.btNumberMask_Click);
            // 
            // btSelectRangeNum
            // 
            this.btSelectRangeNum.Location = new System.Drawing.Point(300, 169);
            this.btSelectRangeNum.Name = "btSelectRangeNum";
            this.btSelectRangeNum.Size = new System.Drawing.Size(28, 23);
            this.btSelectRangeNum.TabIndex = 23;
            this.btSelectRangeNum.Text = "...";
            this.btSelectRangeNum.UseVisualStyleBackColor = true;
            this.btSelectRangeNum.Click += new System.EventHandler(this.btSelectRangeNum_Click);
            // 
            // tbWhereNumber
            // 
            this.tbWhereNumber.Location = new System.Drawing.Point(205, 171);
            this.tbWhereNumber.Name = "tbWhereNumber";
            this.tbWhereNumber.Size = new System.Drawing.Size(82, 20);
            this.tbWhereNumber.TabIndex = 22;
            // 
            // lblWhereNumber
            // 
            this.lblWhereNumber.AutoSize = true;
            this.lblWhereNumber.Location = new System.Drawing.Point(17, 178);
            this.lblWhereNumber.Name = "lblWhereNumber";
            this.lblWhereNumber.Size = new System.Drawing.Size(93, 13);
            this.lblWhereNumber.TabIndex = 21;
            this.lblWhereNumber.Text = "куда разместить";
            // 
            // tbMask
            // 
            this.tbMask.Location = new System.Drawing.Point(78, 104);
            this.tbMask.Name = "tbMask";
            this.tbMask.Size = new System.Drawing.Size(210, 20);
            this.tbMask.TabIndex = 20;
            // 
            // labelTextMask
            // 
            this.labelTextMask.AutoSize = true;
            this.labelTextMask.Location = new System.Drawing.Point(17, 104);
            this.labelTextMask.Name = "labelTextMask";
            this.labelTextMask.Size = new System.Drawing.Size(43, 13);
            this.labelTextMask.TabIndex = 19;
            this.labelTextMask.Text = "Маска:";
            // 
            // labelHeadingMaskSearch
            // 
            this.labelHeadingMaskSearch.AutoSize = true;
            this.labelHeadingMaskSearch.Location = new System.Drawing.Point(17, 72);
            this.labelHeadingMaskSearch.Name = "labelHeadingMaskSearch";
            this.labelHeadingMaskSearch.Size = new System.Drawing.Size(180, 13);
            this.labelHeadingMaskSearch.TabIndex = 18;
            this.labelHeadingMaskSearch.Text = "Выбрать числа из ячеек по маске";
            // 
            // btSelectRange
            // 
            this.btSelectRange.Location = new System.Drawing.Point(300, 22);
            this.btSelectRange.Name = "btSelectRange";
            this.btSelectRange.Size = new System.Drawing.Size(28, 23);
            this.btSelectRange.TabIndex = 17;
            this.btSelectRange.Text = "...";
            this.btSelectRange.UseVisualStyleBackColor = true;
            this.btSelectRange.Click += new System.EventHandler(this.btSelectRange_Click);
            // 
            // tbWhereFIO
            // 
            this.tbWhereFIO.Location = new System.Drawing.Point(205, 24);
            this.tbWhereFIO.Name = "tbWhereFIO";
            this.tbWhereFIO.Size = new System.Drawing.Size(82, 20);
            this.tbWhereFIO.TabIndex = 2;
            // 
            // lblWhereFIO
            // 
            this.lblWhereFIO.AutoSize = true;
            this.lblWhereFIO.Location = new System.Drawing.Point(99, 25);
            this.lblWhereFIO.Name = "lblWhereFIO";
            this.lblWhereFIO.Size = new System.Drawing.Size(93, 13);
            this.lblWhereFIO.TabIndex = 1;
            this.lblWhereFIO.Text = "куда разместить";
            // 
            // btFIO
            // 
            this.btFIO.Location = new System.Drawing.Point(20, 25);
            this.btFIO.Name = "btFIO";
            this.btFIO.Size = new System.Drawing.Size(75, 23);
            this.btFIO.TabIndex = 0;
            this.btFIO.Text = "ФИО";
            this.btFIO.UseVisualStyleBackColor = true;
            this.btFIO.Click += new System.EventHandler(this.btFIO_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.butСommentСell);
            this.tabPage2.Controls.Add(this.tbСommentСell);
            this.tabPage2.Controls.Add(this.labСommentСell);
            this.tabPage2.Controls.Add(this.tbtMaskСomparison);
            this.tabPage2.Controls.Add(this.txMaskSearch);
            this.tabPage2.Controls.Add(this.tbMaskSelection);
            this.tabPage2.Controls.Add(this.label10);
            this.tabPage2.Controls.Add(this.labTextMaskСomparison);
            this.tabPage2.Controls.Add(this.labTextMaskSearch);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.tbSeparator);
            this.tabPage2.Controls.Add(this.labSeparator);
            this.tabPage2.Controls.Add(this.cbTypeSelection);
            this.tabPage2.Controls.Add(this.labTypeSelection);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(346, 419);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Сравнение";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // butСommentСell
            // 
            this.butСommentСell.Location = new System.Drawing.Point(288, 308);
            this.butСommentСell.Name = "butСommentСell";
            this.butСommentСell.Size = new System.Drawing.Size(28, 23);
            this.butСommentСell.TabIndex = 17;
            this.butСommentСell.Text = "...";
            this.butСommentСell.UseVisualStyleBackColor = true;
            // 
            // tbСommentСell
            // 
            this.tbСommentСell.Location = new System.Drawing.Point(216, 310);
            this.tbСommentСell.Name = "tbСommentСell";
            this.tbСommentСell.Size = new System.Drawing.Size(62, 20);
            this.tbСommentСell.TabIndex = 15;
            // 
            // labСommentСell
            // 
            this.labСommentСell.AutoSize = true;
            this.labСommentСell.Location = new System.Drawing.Point(21, 310);
            this.labСommentСell.Name = "labСommentСell";
            this.labСommentСell.Size = new System.Drawing.Size(163, 26);
            this.labСommentСell.TabIndex = 14;
            this.labСommentСell.Text = "Ячейки с которой начинаются \r\nзаполняться комментарии";
            // 
            // tbtMaskСomparison
            // 
            this.tbtMaskСomparison.Location = new System.Drawing.Point(216, 274);
            this.tbtMaskСomparison.Name = "tbtMaskСomparison";
            this.tbtMaskСomparison.Size = new System.Drawing.Size(100, 20);
            this.tbtMaskСomparison.TabIndex = 13;
            // 
            // txMaskSearch
            // 
            this.txMaskSearch.Location = new System.Drawing.Point(216, 238);
            this.txMaskSearch.Name = "txMaskSearch";
            this.txMaskSearch.Size = new System.Drawing.Size(100, 20);
            this.txMaskSearch.TabIndex = 12;
            // 
            // tbMaskSelection
            // 
            this.tbMaskSelection.Location = new System.Drawing.Point(216, 111);
            this.tbMaskSelection.Name = "tbMaskSelection";
            this.tbMaskSelection.Size = new System.Drawing.Size(100, 20);
            this.tbMaskSelection.TabIndex = 11;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(21, 118);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(159, 13);
            this.label10.TabIndex = 10;
            this.label10.Text = "Маска изменения(сравнения)";
            // 
            // labTextMaskСomparison
            // 
            this.labTextMaskСomparison.AutoSize = true;
            this.labTextMaskСomparison.Location = new System.Drawing.Point(21, 281);
            this.labTextMaskСomparison.Name = "labTextMaskСomparison";
            this.labTextMaskСomparison.Size = new System.Drawing.Size(159, 13);
            this.labTextMaskСomparison.TabIndex = 10;
            this.labTextMaskСomparison.Text = "Маска изменения(сравнения)";
            // 
            // labTextMaskSearch
            // 
            this.labTextMaskSearch.AutoSize = true;
            this.labTextMaskSearch.Location = new System.Drawing.Point(21, 241);
            this.labTextMaskSearch.Name = "labTextMaskSearch";
            this.labTextMaskSearch.Size = new System.Drawing.Size(82, 13);
            this.labTextMaskSearch.TabIndex = 9;
            this.labTextMaskSearch.Text = "Маска поиска:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(43, 210);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(190, 13);
            this.label7.TabIndex = 8;
            this.label7.Text = "Строка - функционал не разработан";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(39, 197);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(307, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "Число(#) будем будем приводить к общему виду по маске ";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(22, 171);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(225, 26);
            this.label5.TabIndex = 6;
            this.label5.Text = "Если данные отличаются по оформлению. \r\nПриведем данные к общему виду.";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 142);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Данные В";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Данные А";
            // 
            // tbSeparator
            // 
            this.tbSeparator.Location = new System.Drawing.Point(216, 79);
            this.tbSeparator.Name = "tbSeparator";
            this.tbSeparator.Size = new System.Drawing.Size(100, 20);
            this.tbSeparator.TabIndex = 3;
            // 
            // labSeparator
            // 
            this.labSeparator.AutoSize = true;
            this.labSeparator.Location = new System.Drawing.Point(21, 86);
            this.labSeparator.Name = "labSeparator";
            this.labSeparator.Size = new System.Drawing.Size(73, 13);
            this.labSeparator.TabIndex = 2;
            this.labSeparator.Text = "Разделитель";
            // 
            // cbTypeSelection
            // 
            this.cbTypeSelection.FormattingEnabled = true;
            this.cbTypeSelection.Items.AddRange(new object[] {
            "Число",
            "Строка"});
            this.cbTypeSelection.Location = new System.Drawing.Point(122, 46);
            this.cbTypeSelection.Name = "cbTypeSelection";
            this.cbTypeSelection.Size = new System.Drawing.Size(194, 21);
            this.cbTypeSelection.TabIndex = 1;
            // 
            // labTypeSelection
            // 
            this.labTypeSelection.AutoSize = true;
            this.labTypeSelection.Location = new System.Drawing.Point(21, 54);
            this.labTypeSelection.Name = "labTypeSelection";
            this.labTypeSelection.Size = new System.Drawing.Size(73, 13);
            this.labTypeSelection.TabIndex = 0;
            this.labTypeSelection.Text = "Сравниваем:";
            // 
            // SelectCompareUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tbSelectionCompare);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btRangeTwo);
            this.Controls.Add(this.txtSamplingRangeTwo);
            this.Controls.Add(this.btRangeOne);
            this.Controls.Add(this.txtSamplingRangeOne);
            this.Name = "SelectCompareUserControl";
            this.Size = new System.Drawing.Size(385, 552);
            this.tbSelectionCompare.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btRangeTwo;
        private System.Windows.Forms.TextBox txtSamplingRangeTwo;
        private System.Windows.Forms.Button btRangeOne;
        private System.Windows.Forms.TextBox txtSamplingRangeOne;
        private System.Windows.Forms.TabControl tbSelectionCompare;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label lblWhereFIO;
        private System.Windows.Forms.Button btFIO;
        private System.Windows.Forms.Button btSelectRange;
        private System.Windows.Forms.TextBox tbWhereFIO;
        private System.Windows.Forms.Button btSelectRangeNum;
        private System.Windows.Forms.TextBox tbWhereNumber;
        private System.Windows.Forms.Label lblWhereNumber;
        private System.Windows.Forms.TextBox tbMask;
        private System.Windows.Forms.Label labelTextMask;
        private System.Windows.Forms.Label labelHeadingMaskSearch;
        private System.Windows.Forms.Button btNumberMask;
        private System.Windows.Forms.Label lbTextChange;
        private System.Windows.Forms.TextBox tbChangeMask;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label labTypeSelection;
        private System.Windows.Forms.ComboBox cbTypeSelection;
        private System.Windows.Forms.Label labSeparator;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbSeparator;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label labTextMaskSearch;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label labTextMaskСomparison;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox tbMaskSelection;
        private System.Windows.Forms.Label labСommentСell;
        private System.Windows.Forms.TextBox tbtMaskСomparison;
        private System.Windows.Forms.TextBox txMaskSearch;
        private System.Windows.Forms.Button butСommentСell;
        private System.Windows.Forms.TextBox tbСommentСell;
    }
}
