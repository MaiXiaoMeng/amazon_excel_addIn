using AmazonExcelAddIn.UserControl;
using AmazonExcelAddIn.UserForm;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using MySql.Data.MySqlClient;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;

namespace AmazonExcelAddIn.UserLibrary
{
    public static class CommonHelper
    {
        public static Microsoft.Office.Tools.CustomTaskPane customTaskPane;
        public static string GetRandomSKU()
        {
            char[] Pattern = new char[] {
                '0', '1', '2', '3', '4', '5',
                '6', '7', '8', '9', 'A', 'B',
                'C', 'D', 'E', 'F', 'G', 'H',
                'I', 'J', 'K', 'L', 'M', 'N',
                'O', 'P', 'Q', 'R', 'S', 'T',
                'U', 'V', 'W', 'X', 'Y', 'Z'
            };
            string result = "";
            Random random = new Random(~unchecked((int)DateTime.Now.Ticks));
            for (int i = 0; i < 10; i++)
            {
                if (i == 2 || i == 6) result += "-";
                result += Pattern[random.Next(0, Pattern.Length)];
            }
            return result;

        }
        public static void LoadPlugins()
        {
            DirectoryInfo baseDirectory = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}Plugins/");
            foreach (FileInfo file in baseDirectory.GetFiles())
            {
                if (file.FullName.IndexOf("~$") == -1)
                {
                    switch (file.Extension)
                    {
                        case ".xlam":
                            VariableHelper.Application.Workbooks.Open(file.FullName);
                            break;
                        case ".xll":
                            VariableHelper.Application.RegisterXLL(file.FullName);
                            break;
                    }
                }
            }
        }
        public static string MD5(string encryptString)
        {
            byte[] result = Encoding.Default.GetBytes(encryptString);
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] output = md5.ComputeHash(result);
            string encryptResult = BitConverter.ToString(output).Replace("-", "");
            return encryptResult;
        }
        public static void OpenNavigationBar()
        {
            if (customTaskPane == null)
            {
                NavigationBar navigationBar = new NavigationBar();
                customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(navigationBar, "导航栏");
                customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                customTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                customTaskPane.Width = 200;
            }
            customTaskPane.Visible = true;
        }
        public static void RibbonMainLoad()
        {
            Globals.Ribbons.RibbonMain.BarcodeWidth.Text = VariableHelper.BarcodeWidth;
            Globals.Ribbons.RibbonMain.BarcodeHeight.Text = VariableHelper.BarcodeHeight;
            Globals.Ribbons.RibbonMain.BarcodeTop.Text = VariableHelper.BarcodeTop;
            Globals.Ribbons.RibbonMain.BarcodeBottom.Text = VariableHelper.BarcodeBottom;
            Globals.Ribbons.RibbonMain.BarcodeLeft.Text = VariableHelper.BarcodeLeft;
            Globals.Ribbons.RibbonMain.BarcodeRight.Text = VariableHelper.BarcodeRight;
            Globals.Ribbons.RibbonMain.BarcodePrinterName.Text = VariableHelper.BarcodePrinterName;
        }

    }
}
