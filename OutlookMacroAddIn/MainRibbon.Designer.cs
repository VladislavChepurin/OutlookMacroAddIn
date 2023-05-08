namespace OutlookMacroAddIn
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Абиэлт";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Основные макросы";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.label1);
            this.group2.Items.Add(this.label2);
            this.group2.Items.Add(this.label3);
            this.group2.Label = "Курсы валют ЦБ РФ";
            this.group2.Name = "group2";
            // 
            // label1
            // 
            this.label1.Label = "label1";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "label2";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "label3";
            this.label3.Name = "label3";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button2);
            this.group3.Items.Add(this.button3);
            this.group3.Name = "group3";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::OutlookMacroAddIn.Resource._101;
            this.button1.Label = "Преобразовать в проект";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::OutlookMacroAddIn.Properties.Resources._201;
            this.button2.Label = "About";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::OutlookMacroAddIn.Properties.Resources._202;
            this.button3.Label = "Открыть папку";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
