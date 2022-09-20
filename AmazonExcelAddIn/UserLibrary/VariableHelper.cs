using Excel = Microsoft.Office.Interop.Excel;
using Ribbon = Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace AmazonExcelAddIn.UserLibrary
{
    public class VariableHelper
    {
        public static Excel.Application Application
        {
            get {
                try
                {
                    return Globals.ThisAddIn.Application;
                }
                catch (Exception)
                {
                    return (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
            }
        
        }

        //public static RibbonMain RibbonMain
        //{
        //    get
        //    {
        //        try
        //        {
        //            return Globals.Ribbons.Base;
        //        }
        //        catch (Exception)
        //        {
        //            return (RibbonMain)Marshal.GetActiveObject("Excel.RibbonMain");
        //        }
        //    }

        //}

        public static string BarcodeBottom { get { return ConfigHelper.Get("BarcodeBottom", Default: "15"); }}
        public static string BarcodeHeight { get { return ConfigHelper.Get("BarcodeHeight", Default: "282"); }}
        public static string BarcodeLeft { get { return ConfigHelper.Get("BarcodeLeft", Default: "20"); }}
        public static string BarcodeRight { get { return ConfigHelper.Get("BarcodeRight", Default: "20"); }}
        public static string BarcodeTop { get { return ConfigHelper.Get("BarcodeTop", Default: "15"); }}
        public static string BarcodeWidth { get { return ConfigHelper.Get("BarcodeWidth", Default: "689"); }}
        public static string BarcodePrinterName { get { return ConfigHelper.Get("BarcodePrinterName", Default: ""); }}
        public static ListObjectHelper ListObject { get { return new ListObjectHelper("配置-文件配置", "视图_数据库_配置信息"); }}
        public static string MysqlDatabase { get { return ListObject.GetValue(0, "数据库名称"); } }
        public static string DefaultWarehouse { get { return ListObject.GetValue(0, "默认仓库"); } }
        public static string MysqlHost { get { return ListObject.GetValue(0, "数据库地址"); } }
        public static string MysqlPassWord { get { return ListObject.GetValue(0, "数据库密码"); } }
        public static string MysqlUserName { get { return ListObject.GetValue(0, "数据库用户"); } }
        public static string PrincipalName { get { return ListObject.GetValue(0, "用户名"); } }
        public static MySQLHelper MySql { get { return new MySQLHelper(Default: true); } }
 
    }
}
