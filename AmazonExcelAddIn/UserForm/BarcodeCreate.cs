using AmazonExcelAddIn.UserLibrary;
using Spire.Pdf;
using Spire.Pdf.Barcode;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace AmazonExcelAddIn.UserForm
{
    public partial class BarcodeCreate : Form
    {
        public string fnSKU;
        public BarcodeCreate(string fnSKU="")
        {
            InitializeComponent();
            this.fnSKU = fnSKU;
        }

        private void BarcodeCreate_Load(object sender, EventArgs e)
        {
            labelText.Text = SpireHelper.GetLabel(fnSKU);
            PreviewBarcode.Image = SpireHelper.CreateBarcode(fnSKU).SaveAsImage(0, 30, 30);
        }

        private void RefreshBarcode_Click(object sender, EventArgs e)
        {
            PreviewBarcode.Image = SpireHelper.CreateBarcode(fnSKU, labelText.Text).SaveAsImage(0, 30, 30);
        }

        private void BarcodeSave_Click(object sender, EventArgs e)
        {
            DataRow barcodeInformation = VariableHelper.MySql.ExecuteDataRow($"SELECT FNSKU,`短名称` FROM `库存管理_条码信息` WHERE FNSKU = '{fnSKU}'");
            if (barcodeInformation == null)
            {
                VariableHelper.MySql.ExecuteNonQuery($"INSERT INTO `库存管理_条码信息`(`FNSKU`, `短名称`) VALUES ('{fnSKU}', '{labelText.Text}')");
            }
            else
            {
                VariableHelper.MySql.ExecuteNonQuery($"UPDATE `库存管理_条码信息` SET `短名称` = '{labelText.Text}' WHERE `FNSKU` = '{fnSKU}'");
            }
            Dispose();
            Close();
        }
    }
}
