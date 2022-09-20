
namespace AmazonExcelAddIn.UserControl
{
    partial class NavigationBar
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NavigationBar));
            this.TVMenu = new HZH_Controls.Controls.TreeViewEx();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TVMenu
            // 
            this.TVMenu.BackColor = System.Drawing.SystemColors.ControlLight;
            this.TVMenu.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TVMenu.DrawMode = System.Windows.Forms.TreeViewDrawMode.OwnerDrawAll;
            this.TVMenu.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TVMenu.ForeColor = System.Drawing.Color.Transparent;
            this.TVMenu.FullRowSelect = true;
            this.TVMenu.HideSelection = false;
            this.TVMenu.IsShowByCustomModel = true;
            this.TVMenu.IsShowTip = false;
            this.TVMenu.ItemHeight = 50;
            this.TVMenu.Location = new System.Drawing.Point(3, 3);
            this.TVMenu.LstTips = ((System.Collections.Generic.Dictionary<string, string>)(resources.GetObject("TVMenu.LstTips")));
            this.TVMenu.Name = "TVMenu";
            this.TVMenu.NodeBackgroundColor = System.Drawing.SystemColors.Control;
            this.TVMenu.NodeDownPic = ((System.Drawing.Image)(resources.GetObject("TVMenu.NodeDownPic")));
            this.TVMenu.NodeForeColor = System.Drawing.Color.Green;
            this.TVMenu.NodeHeight = 50;
            this.TVMenu.NodeIsShowSplitLine = true;
            this.TVMenu.NodeSelectedColor = System.Drawing.SystemColors.AppWorkspace;
            this.TVMenu.NodeSelectedForeColor = System.Drawing.Color.White;
            this.TVMenu.NodeSplitLineColor = System.Drawing.Color.FromArgb(((int)(((byte)(232)))), ((int)(((byte)(232)))), ((int)(((byte)(232)))));
            this.TVMenu.NodeUpPic = ((System.Drawing.Image)(resources.GetObject("TVMenu.NodeUpPic")));
            this.TVMenu.ParentNodeCanSelect = true;
            this.TVMenu.ShowLines = false;
            this.TVMenu.ShowPlusMinus = false;
            this.TVMenu.ShowRootLines = false;
            this.TVMenu.Size = new System.Drawing.Size(186, 450);
            this.TVMenu.TabIndex = 0;
            this.TVMenu.TipFont = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.TVMenu.TipImage = ((System.Drawing.Image)(resources.GetObject("TVMenu.TipImage")));
            this.TVMenu.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.TVMenu_AfterSelect);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.TVMenu);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(192, 456);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // NavigationBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Name = "NavigationBar";
            this.Size = new System.Drawing.Size(192, 456);
            this.Load += new System.EventHandler(this.NavigationBar_Load);
            this.SizeChanged += new System.EventHandler(this.NavigationBar_SizeChanged);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private HZH_Controls.Controls.TreeViewEx TVMenu;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
    }
}
