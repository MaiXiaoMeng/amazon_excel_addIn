using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using AmazonExcelAddIn.UserLibrary;
using AmazonExcelAddIn.UserForm;
using Microsoft.VisualBasic;
using System.Drawing.Printing;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.Reflection;

namespace AmazonExcelAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// 激活工作表的事件
        /// </summary>
        /// <param name="Sh">被激活的工作表对象</param>
        private void Application_SheetActivate(object Sh)
        {
            LoggerHelper.Debug($"执行 RibbonMain 刷新");
            CommonHelper.RibbonMainLoad();
            LoggerHelper.Debug($"把 object 对象的 Sh 变量强行转换成 Excel.Worksheet 对象");
            Excel.Worksheet sh = (Excel.Worksheet)Sh;
            Globals.Ribbons.RibbonMain.BarcodePrinterConfig.Visible = sh.Name == "仓库-出库计划";
        }
        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            ListObjectHelper listObjectHelper = new ListObjectHelper("产品-产品菜单", "视图_库存管理_菜单管理_列表菜单");
            LoggerHelper.Debug($"变量 listObject > 产品-产品菜单>>视图_库存管理_菜单管理_列表菜单");

            LoggerHelper.Debug($"遍历 listObject > Count: {listObjectHelper.Source.ListRows.Count}");
            for (int index = 0; index < listObjectHelper.Source.ListRows.Count; index++)
            {
                string sheetName = listObjectHelper.GetValue(index, "Sheet");
                LoggerHelper.Debug($"变量 sheetName > {sheetName}");
                string listObjectName = listObjectHelper.GetValue(index, "ListObject");
                LoggerHelper.Debug($"变量 listObjectName > {listObjectName}");
                string menuName = listObjectHelper.GetValue(index, "Menu");
                LoggerHelper.Debug($"变量 menuName > {menuName}");
                string fieldName = listObjectHelper.GetValue(index, "Field");
                LoggerHelper.Debug($"变量 fieldName > {fieldName}");
                string minName = listObjectHelper.GetValue(index, "Min");
                LoggerHelper.Debug($"变量 minName > {minName}");
                string maxName = listObjectHelper.GetValue(index, "Max");
                LoggerHelper.Debug($"变量 maxName > {maxName}");

                bool isSheetName = ((Excel.Worksheet)Sh).Name == sheetName;
                LoggerHelper.Debug($"判断 右键点击 的工作簿名字是否等于当前变量的名字 > {isSheetName}");
                if (isSheetName)
                {
                    int ColInd = new ListObjectHelper(sheetName, listObjectName).Source.ListColumns.get_Item(fieldName).Index;
                    LoggerHelper.Debug($"变量 ColInd > {ColInd}");

                    if (Target.Count >= Convert.ToInt32(minName) && Target.Count <= Convert.ToInt32(maxName) && Target.Column == ColInd)
                    {
                        try
                        {
                            LoggerHelper.Debug($"执行 右键菜单{menuName} {fieldName}显示");
                            VariableHelper.Application.CommandBars[menuName].ShowPopup();
                            LoggerHelper.Debug($"取消 原生右键菜单显示");
                            Cancel = true;
                        }
                        catch (Exception error)
                        {
                            LoggerHelper.Error($"执行 右键菜单{menuName} 显示出错");
                            LoggerHelper.Error(error.ToString());
                            VariableHelper.Application.StatusBar = $"调用菜单出错,请刷新菜单数据后再尝试!";
                        }
                    }
                }
            }
        }
        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            bool isWorkbookName = Wb.Name == $"库存管理.xlsm";
            LoggerHelper.Debug($"判断 工作簿 是不是 库存管理.xlsm > {isWorkbookName}");
            if (isWorkbookName)
            {
                LoggerHelper.Debug($"执行 打开左侧的导航栏");
                CommonHelper.OpenNavigationBar();

                LoggerHelper.Debug($"绑定 Application_SheetActivate 事件");
                VariableHelper.Application.SheetActivate += Application_SheetActivate;
                LoggerHelper.Debug($"绑定 Application_SheetBeforeRightClick 事件");
                VariableHelper.Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;

                LoggerHelper.Debug($"激活 首页 Sheet");
                VariableHelper.Application.ActiveWorkbook.Sheets["首页"].Activate();

                bool isPrincipalName = VariableHelper.Application.ActiveWorkbook.Sheets["配置-文件配置"].Cells.Range["A2"].Value == null;
                LoggerHelper.Debug($"判断 负责人 是否填写 > {isPrincipalName}");
                if (isPrincipalName)
                {
                    string principalName = Interaction.InputBox("请输入你的名字");
                    LoggerHelper.Debug($"填写 负责人 > {principalName}");
                    VariableHelper.Application.ActiveWorkbook.Sheets["配置-文件配置"].Cells.Range["A2"].Value = principalName;

                    LoggerHelper.Debug($"保存 工作簿");
                    VariableHelper.Application.ActiveWorkbook.Save();
                }

                LoggerHelper.Debug($"刷新 右键菜单");
                ContextMenuHelper.RefreshMenu();
            }
            bool isDataManagement =  Wb.Name == "库存管理.xlsm";
            LoggerHelper.Debug($"激活 TAB 的 数据管理 分组 > {isDataManagement}");
            Globals.Ribbons.RibbonMain.DataManagement.Visible = isDataManagement;


        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            LoggerHelper.Debug($"绑定 Application_WorkbookOpen 事件");
            VariableHelper.Application.WorkbookOpen += Application_WorkbookOpen;
            LoggerHelper.Debug($"加载第三方插件");
            CommonHelper.LoadPlugins();
        }
        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
