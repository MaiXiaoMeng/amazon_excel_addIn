using AmazonExcelAddIn;
using AmazonExcelAddIn.UserControl;
using AmazonExcelAddIn.UserForm;
using AmazonExcelAddIn.UserLibrary;
using AmazonExcelUDF.ExcelUDF;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace AmazonExcelAddInTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        

            //foreach (CommandBar commandBar in VariableHelper.Application.CommandBars)
            //{
            //    try
            //    {
            //        Console.WriteLine(commandBar.Name);
            //        commandBar.ShowPopup();
            //    }
            //    catch (Exception)
            //    {

            //    }

            //}
        }
    }
}
