
namespace AmazonExcelAddIn
{
    partial class RibbonMain : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMain()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.InventoryManagement = this.Factory.CreateRibbonTab();
            this.DataManagement = this.Factory.CreateRibbonGroup();
            this.DataRefresh = this.Factory.CreateRibbonButton();
            this.RefreshMenu = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.CommonOperations = this.Factory.CreateRibbonMenu();
            this.RandomSKU = this.Factory.CreateRibbonButton();
            this.BarcodePrinterConfig = this.Factory.CreateRibbonGroup();
            this.BarcodeLayoutBox1 = this.Factory.CreateRibbonBox();
            this.BarcodePrinterName = this.Factory.CreateRibbonComboBox();
            this.BarcodeLayoutBox2 = this.Factory.CreateRibbonBox();
            this.BarcodeHeight = this.Factory.CreateRibbonEditBox();
            this.BarcodeWidth = this.Factory.CreateRibbonEditBox();
            this.BarcodeLayoutBox3 = this.Factory.CreateRibbonBox();
            this.BarcodeTop = this.Factory.CreateRibbonEditBox();
            this.BarcodeBottom = this.Factory.CreateRibbonEditBox();
            this.BarcodeLeft = this.Factory.CreateRibbonEditBox();
            this.BarcodeRight = this.Factory.CreateRibbonEditBox();
            this.InventoryManagement.SuspendLayout();
            this.DataManagement.SuspendLayout();
            this.group1.SuspendLayout();
            this.BarcodePrinterConfig.SuspendLayout();
            this.BarcodeLayoutBox1.SuspendLayout();
            this.BarcodeLayoutBox2.SuspendLayout();
            this.BarcodeLayoutBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // InventoryManagement
            // 
            this.InventoryManagement.Groups.Add(this.DataManagement);
            this.InventoryManagement.Groups.Add(this.group1);
            this.InventoryManagement.Groups.Add(this.BarcodePrinterConfig);
            this.InventoryManagement.Label = "MaiXiaoMeng";
            this.InventoryManagement.Name = "InventoryManagement";
            // 
            // DataManagement
            // 
            this.DataManagement.Items.Add(this.DataRefresh);
            this.DataManagement.Items.Add(this.RefreshMenu);
            this.DataManagement.Label = "数据操作";
            this.DataManagement.Name = "DataManagement";
            // 
            // DataRefresh
            // 
            this.DataRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DataRefresh.Label = "                     数据更新";
            this.DataRefresh.Name = "DataRefresh";
            this.DataRefresh.OfficeImageId = "DatabaseLinedTableManager";
            this.DataRefresh.ShowImage = true;
            this.DataRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RefreshData_Click);
            // 
            // RefreshMenu
            // 
            this.RefreshMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RefreshMenu.Label = "                     刷新菜单";
            this.RefreshMenu.Name = "RefreshMenu";
            this.RefreshMenu.OfficeImageId = "TablePropertiesInfoPath";
            this.RefreshMenu.ShowImage = true;
            this.RefreshMenu.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RefreshMenu_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.CommonOperations);
            this.group1.Label = "常用操作";
            this.group1.Name = "group1";
            // 
            // CommonOperations
            // 
            this.CommonOperations.Items.Add(this.RandomSKU);
            this.CommonOperations.Label = "常用操作";
            this.CommonOperations.Name = "CommonOperations";
            // 
            // RandomSKU
            // 
            this.RandomSKU.Label = "填充随机SKU";
            this.RandomSKU.Name = "RandomSKU";
            this.RandomSKU.ShowImage = true;
            this.RandomSKU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RandomSKU_Click);
            // 
            // BarcodePrinterConfig
            // 
            this.BarcodePrinterConfig.Items.Add(this.BarcodeLayoutBox1);
            this.BarcodePrinterConfig.Items.Add(this.BarcodeLayoutBox2);
            this.BarcodePrinterConfig.Items.Add(this.BarcodeLayoutBox3);
            this.BarcodePrinterConfig.Label = "条码打印机设置";
            this.BarcodePrinterConfig.Name = "BarcodePrinterConfig";
            this.BarcodePrinterConfig.Visible = false;
            // 
            // BarcodeLayoutBox1
            // 
            this.BarcodeLayoutBox1.Items.Add(this.BarcodePrinterName);
            this.BarcodeLayoutBox1.Name = "BarcodeLayoutBox1";
            // 
            // BarcodePrinterName
            // 
            this.BarcodePrinterName.Label = "条码打印机";
            this.BarcodePrinterName.Name = "BarcodePrinterName";
            this.BarcodePrinterName.SizeString = "1234567890123456789012345";
            this.BarcodePrinterName.Text = null;
            this.BarcodePrinterName.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BarcodePrinterName_TextChanged);
            // 
            // BarcodeLayoutBox2
            // 
            this.BarcodeLayoutBox2.Items.Add(this.BarcodeHeight);
            this.BarcodeLayoutBox2.Items.Add(this.BarcodeWidth);
            this.BarcodeLayoutBox2.Name = "BarcodeLayoutBox2";
            // 
            // BarcodeHeight
            // 
            this.BarcodeHeight.Label = "条码大小高";
            this.BarcodeHeight.Name = "BarcodeHeight";
            this.BarcodeHeight.SizeString = "12345678";
            this.BarcodeHeight.Text = null;
            this.BarcodeHeight.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditBox_TextChanged);
            // 
            // BarcodeWidth
            // 
            this.BarcodeWidth.Label = "条码大小宽";
            this.BarcodeWidth.Name = "BarcodeWidth";
            this.BarcodeWidth.SizeString = "12345678";
            this.BarcodeWidth.Text = null;
            this.BarcodeWidth.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditBox_TextChanged);
            // 
            // BarcodeLayoutBox3
            // 
            this.BarcodeLayoutBox3.Items.Add(this.BarcodeTop);
            this.BarcodeLayoutBox3.Items.Add(this.BarcodeBottom);
            this.BarcodeLayoutBox3.Items.Add(this.BarcodeLeft);
            this.BarcodeLayoutBox3.Items.Add(this.BarcodeRight);
            this.BarcodeLayoutBox3.Name = "BarcodeLayoutBox3";
            // 
            // BarcodeTop
            // 
            this.BarcodeTop.Label = "上边距";
            this.BarcodeTop.Name = "BarcodeTop";
            this.BarcodeTop.SizeString = "00";
            this.BarcodeTop.Text = null;
            this.BarcodeTop.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditBox_TextChanged);
            // 
            // BarcodeBottom
            // 
            this.BarcodeBottom.Label = "下边距";
            this.BarcodeBottom.Name = "BarcodeBottom";
            this.BarcodeBottom.SizeString = "00";
            this.BarcodeBottom.Text = null;
            this.BarcodeBottom.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditBox_TextChanged);
            // 
            // BarcodeLeft
            // 
            this.BarcodeLeft.Label = "左边距";
            this.BarcodeLeft.Name = "BarcodeLeft";
            this.BarcodeLeft.SizeString = "00";
            this.BarcodeLeft.Text = null;
            this.BarcodeLeft.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditBox_TextChanged);
            // 
            // BarcodeRight
            // 
            this.BarcodeRight.Label = "右边距";
            this.BarcodeRight.Name = "BarcodeRight";
            this.BarcodeRight.SizeString = "00";
            this.BarcodeRight.Text = null;
            this.BarcodeRight.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditBox_TextChanged);
            // 
            // RibbonMain
            // 
            this.Name = "RibbonMain";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.InventoryManagement);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMain_Load);
            this.InventoryManagement.ResumeLayout(false);
            this.InventoryManagement.PerformLayout();
            this.DataManagement.ResumeLayout(false);
            this.DataManagement.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.BarcodePrinterConfig.ResumeLayout(false);
            this.BarcodePrinterConfig.PerformLayout();
            this.BarcodeLayoutBox1.ResumeLayout(false);
            this.BarcodeLayoutBox1.PerformLayout();
            this.BarcodeLayoutBox2.ResumeLayout(false);
            this.BarcodeLayoutBox2.PerformLayout();
            this.BarcodeLayoutBox3.ResumeLayout(false);
            this.BarcodeLayoutBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab InventoryManagement;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup BarcodePrinterConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox BarcodeLayoutBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox BarcodePrinterName;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox BarcodeLayoutBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox BarcodeHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox BarcodeWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox BarcodeLayoutBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox BarcodeTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox BarcodeBottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox BarcodeLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox BarcodeRight;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup DataManagement;
        public Microsoft.Office.Tools.Ribbon.RibbonButton DataRefresh;
        public Microsoft.Office.Tools.Ribbon.RibbonButton RefreshMenu;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonMenu CommonOperations;
        public Microsoft.Office.Tools.Ribbon.RibbonButton RandomSKU;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMain RibbonMain
        {
            get { return this.GetRibbon<RibbonMain>(); }
        }
    }
}
