using Spire.Pdf;
using Spire.Pdf.Barcode;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AmazonExcelAddIn.UserLibrary
{
    public static class SpireHelper
    {
        public static PdfDocument CreateBarcode(string fnSKU = "Null", string labelText = "")
        {
            PdfDocument pdfDocument = new PdfDocument();
            string labelSource = "Made in China    ";
            string labelTitle = GetTitle(fnSKU);
            if (labelText == "") labelText = GetLabel(fnSKU);
            if (labelText != "" && labelTitle != "")
            {
                float drawLastHeight;
                PdfLayoutResult drawLayoutResult;

                // PDF文档_页面.Canvas.DrawImage(条码图片, 0, 0, 条码大小_宽, 条码图片.Height);
                // PDF文档.SaveToFile("MyFirstPDF.pdf");
                // PDF文档.LoadFromFile("MyFirstPDF.pdf"); // 加载 PDF 文件
                // PdfImage 条码图片 = PdfImage.FromFile("1.png");
                // int 条码大小_宽 = 条码图片.Width;
                // int 条码大小_高 = 条码图片.Height;

                float barcodeWidth = float.Parse(VariableHelper.BarcodeWidth);
                float barcodeHeight = float.Parse(VariableHelper.BarcodeHeight);
                float barcodeTop = float.Parse(VariableHelper.BarcodeTop);
                float barcodeBottom = float.Parse(VariableHelper.BarcodeBottom);
                float barcodeLeft = float.Parse(VariableHelper.BarcodeLeft);
                float barcodeRight = float.Parse(VariableHelper.BarcodeRight);

                SizeF barcodeSize = new SizeF(barcodeWidth, barcodeHeight);
                PdfMargins barcodeMargin = new PdfMargins(barcodeLeft, barcodeTop, barcodeRight, barcodeBottom);
                PdfPageBase pdfDocumentPageBase = pdfDocument.Pages.Add(barcodeSize, barcodeMargin);

                PdfCode128ABarcode productBarcode = new PdfCode128ABarcode(fnSKU)
                {
                    BarHeight = 110,
                    NarrowBarWidth = 4.5f,
                    BarcodeToTextGapHeight = 1f,
                    TextDisplayLocation = TextLocation.Bottom,
                    TextColor = Color.Black,
                    Font = new PdfTrueTypeFont(new Font("Arial", 50f, FontStyle.Bold), true)
                };
                productBarcode.Draw(pdfDocumentPageBase, new PointF(0, 0));
                drawLastHeight = productBarcode.Bounds.Bottom;

                PdfTextWidget productTitleText = new PdfTextWidget
                {
                    Font = new PdfTrueTypeFont(new Font("Arial", 35f, FontStyle.Regular), true),
                    Text = labelTitle
                };
                drawLayoutResult = productTitleText.Draw(pdfDocumentPageBase, 0, drawLastHeight + 40);
                drawLastHeight = drawLayoutResult.Bounds.Bottom;

                PdfTextWidget labelSourceText = new PdfTextWidget
                {
                    Font = new PdfTrueTypeFont(new Font("Gabriola", 45f, FontStyle.Bold), true),
                    Text = labelSource
                };
                labelSourceText.Draw(pdfDocumentPageBase, 1, drawLastHeight - 26);
                drawLayoutResult = labelSourceText.Draw(pdfDocumentPageBase, 0, drawLastHeight - 25);

                PdfTextWidget productInformationText = new PdfTextWidget
                {
                    Font = new PdfTrueTypeFont(new Font("黑体", 36f, FontStyle.Bold), true),
                    Text = labelText
                };
                productInformationText.Draw(pdfDocumentPageBase, drawLayoutResult.Bounds.Right + 10, drawLastHeight);
            }
            pdfDocument.SaveToFile($"productBarcode/{fnSKU}.pdf");
            return pdfDocument;
        }
        public static string GetLabel(string fnSKU)
        {
            DataRow barcodeInformation = VariableHelper.MySql.ExecuteDataRow($"SELECT FNSKU,`短名称` FROM `库存管理_条码信息` WHERE FNSKU = '{fnSKU}'");
            if (barcodeInformation != null)
            {
                return barcodeInformation["短名称"].ToString();
            }
            else
            {
                DataRow productInformation = VariableHelper.MySql.ExecuteDataRow($"SELECT `产品名称`,`产品型号`,`产品颜色` FROM `视图_库存管理_产品信息_sku信息` WHERE FNSKU = '{fnSKU}'");
                if (productInformation == null)
                {
                    return $"条码:{fnSKU} 对应的信息没填完整, 请填好在尝试制作条码!";
                }
                else
                {
                    string productName = productInformation["产品名称"].ToString();
                    string productModel = productInformation["产品型号"].ToString();
                    string productColor = productInformation["产品颜色"].ToString();
                    return $"{productName}|{productModel}|{productColor}";
                }
            }
        }
        public static PaperSize GetPaperSize(string paperName, float barcodeWidth, float barcodeHeight)
        {
            PrintDocument printDocument = new PrintDocument();
            printDocument.PrinterSettings.PrinterName = VariableHelper.BarcodePrinterName;
            foreach (PaperSize paperSize in printDocument.PrinterSettings.PaperSizes)
            {
                if (paperSize.PaperName == paperName)
                {
                    return paperSize;
                }
            }
            return new PaperSize(paperName, (int)((barcodeWidth / 2.54) * 100), (int)((barcodeHeight / 2.54) * 100));
        }
        public static string GetTitle(string fnSKU)
        {
            DataRow productInformation = VariableHelper.MySql.ExecuteDataRow($"SELECT `商品名称` FROM `库存报告_管理亚马逊库存` WHERE FNSKU = '{fnSKU}'");
            if (productInformation != null)
            {
                int labelMaxLength = 35;
                string modelColor;
                string productTitle = productInformation["商品名称"].ToString();
                int positionLeftParenthesis = productTitle.IndexOf("(");
                int positionRightParenthesis = productTitle.IndexOf(")");
                if (positionLeftParenthesis > -1 && positionRightParenthesis > -1)
                {
                    modelColor = productTitle.Substring(positionLeftParenthesis, positionRightParenthesis - positionLeftParenthesis + 1);
                }
                else if (productTitle.IndexOf("-") > -1)
                {
                    modelColor = productTitle.Split('-').Last();
                }
                else
                {
                    modelColor = productTitle.Split(' ').Last();
                }
                string labelLeftText = productTitle.Substring(0, labelMaxLength - modelColor.Length);
                return $"{labelLeftText}...{modelColor}";
            }
            else
            {
                return "";
            }
        }
        public static void PrintBarcode(string fnSKU, short printNumber)
        {
            float barcodeWidth = float.Parse(VariableHelper.BarcodeWidth);
            float barcodeHeight = float.Parse(VariableHelper.BarcodeHeight);
            PdfDocument barcodeDocument = CreateBarcode(fnSKU);
            barcodeDocument.PrintSettings.PrinterName = VariableHelper.BarcodePrinterName;
            barcodeDocument.PrintSettings.Copies = printNumber;
            barcodeDocument.PrintSettings.Collate = true;
            barcodeDocument.PrintSettings.PaperSize = GetPaperSize("亚马逊产品标", barcodeWidth, barcodeHeight);
            barcodeDocument.Print();
            barcodeDocument.Dispose();
            barcodeDocument.Close();
        }
    }
}
