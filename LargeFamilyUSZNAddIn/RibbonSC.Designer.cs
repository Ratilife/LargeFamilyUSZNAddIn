namespace LargeFamilyUSZNAddIn
{
    partial class RibbonSC : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonSC()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupSelectCompare = this.Factory.CreateRibbonGroup();
            this.butSC = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupSelectCompare.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupSelectCompare);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupSelectCompare
            // 
            this.groupSelectCompare.Items.Add(this.butSC);
            this.groupSelectCompare.Name = "groupSelectCompare";
            // 
            // butSC
            // 
            this.butSC.Label = "Выборка/Сравнение";
            this.butSC.Name = "butSC";
            this.butSC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butSC_Click);
            // 
            // RibbonSC
            // 
            this.Name = "RibbonSC";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonSC_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupSelectCompare.ResumeLayout(false);
            this.groupSelectCompare.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSelectCompare;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butSC;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonSC RibbonSC
        {
            get { return this.GetRibbon<RibbonSC>(); }
        }
    }
}
