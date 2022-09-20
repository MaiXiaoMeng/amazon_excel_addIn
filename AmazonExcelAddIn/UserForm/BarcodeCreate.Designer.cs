
namespace AmazonExcelAddIn.UserForm
{
    partial class BarcodeCreate
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.PreviewBarcode = new System.Windows.Forms.PictureBox();
            this.labelText = new System.Windows.Forms.TextBox();
            this.RefreshBarcode = new System.Windows.Forms.Button();
            this.BarcodeSave = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.PreviewBarcode)).BeginInit();
            this.SuspendLayout();
            // 
            // PreviewBarcode
            // 
            this.PreviewBarcode.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PreviewBarcode.Location = new System.Drawing.Point(0, 29);
            this.PreviewBarcode.Name = "PreviewBarcode";
            this.PreviewBarcode.Size = new System.Drawing.Size(317, 129);
            this.PreviewBarcode.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.PreviewBarcode.TabIndex = 0;
            this.PreviewBarcode.TabStop = false;
            // 
            // labelText
            // 
            this.labelText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelText.Location = new System.Drawing.Point(12, 3);
            this.labelText.Name = "labelText";
            this.labelText.Size = new System.Drawing.Size(179, 21);
            this.labelText.TabIndex = 2;
            // 
            // RefreshBarcode
            // 
            this.RefreshBarcode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.RefreshBarcode.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.RefreshBarcode.Location = new System.Drawing.Point(197, 2);
            this.RefreshBarcode.Name = "RefreshBarcode";
            this.RefreshBarcode.Size = new System.Drawing.Size(54, 23);
            this.RefreshBarcode.TabIndex = 3;
            this.RefreshBarcode.Text = "刷新";
            this.RefreshBarcode.UseVisualStyleBackColor = true;
            this.RefreshBarcode.Click += new System.EventHandler(this.RefreshBarcode_Click);
            // 
            // BarcodeSave
            // 
            this.BarcodeSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BarcodeSave.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BarcodeSave.Location = new System.Drawing.Point(257, 2);
            this.BarcodeSave.Name = "BarcodeSave";
            this.BarcodeSave.Size = new System.Drawing.Size(54, 23);
            this.BarcodeSave.TabIndex = 4;
            this.BarcodeSave.Text = "保存";
            this.BarcodeSave.UseVisualStyleBackColor = true;
            this.BarcodeSave.Click += new System.EventHandler(this.BarcodeSave_Click);
            // 
            // BarcodeCreate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(317, 160);
            this.Controls.Add(this.BarcodeSave);
            this.Controls.Add(this.RefreshBarcode);
            this.Controls.Add(this.labelText);
            this.Controls.Add(this.PreviewBarcode);
            this.Name = "BarcodeCreate";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BarcodeCreate";
            this.Load += new System.EventHandler(this.BarcodeCreate_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PreviewBarcode)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox PreviewBarcode;
        private System.Windows.Forms.TextBox labelText;
        private System.Windows.Forms.Button RefreshBarcode;
        private System.Windows.Forms.Button BarcodeSave;
    }
}