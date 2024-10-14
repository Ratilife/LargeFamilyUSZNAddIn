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
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tbSelectionCompare.SuspendLayout();
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
            // 
            // txtSamplingRangeTwo
            // 
            this.txtSamplingRangeTwo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
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
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(346, 419);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Выборка";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(346, 419);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Сравнение";
            this.tabPage2.UseVisualStyleBackColor = true;
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
    }
}
