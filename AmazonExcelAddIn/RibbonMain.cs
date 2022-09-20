
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using AmazonExcelAddIn.UserControl;
using AmazonExcelAddIn.UserLibrary;
using System.Diagnostics;
using System.Drawing.Printing;
using Microsoft.Office.Interop.Excel;

namespace AmazonExcelAddIn
{
    public partial class RibbonMain
    {
        public void BarcodePrinterName_TextChanged(object sender = null, RibbonControlEventArgs e = null)
        {
            ConfigHelper.Set(((RibbonComboBox)sender).Name, ((RibbonComboBox)sender).Text);
        }
        public void RandomSKU_Click(object sender = null, RibbonControlEventArgs e = null)
        {
            Range selectionCellRange = VariableHelper.Application.Selection;
            if (selectionCellRange != null)
            {
                foreach (Range range in selectionCellRange.Rows)
                {
                    range.Value = CommonHelper.GetRandomSKU();
                }
                VariableHelper.Application.StatusBar = $"随机SKU填充完成";
            }
        }
        public void RefreshData_Click(object sender = null, RibbonControlEventArgs e = null)
        {
            VariableHelper.Application.StatusBar = $"正在更新数据";
            VariableHelper.Application.ScreenUpdating = false;
            string activeSheet = VariableHelper.Application.ActiveSheet.Name;
            foreach (Excel.ListObject listObject in VariableHelper.Application.ActiveSheet.ListObjects)
            {
                ListObjectHelper listObjectHelper = new ListObjectHelper(activeSheet, listObject.Name);
                for (int index = 0; index < listObjectHelper.Source.ListRows.Count; index++)
                {
                    string databaseTableName = "";
                    string id = "";
                    switch (listObjectHelper.GetValue(index, "操作类型", false))
                    {
                        case "新增":
                        case "修改":
                        case "删除":
                            switch (activeSheet)
                            {
                                #region 配置-文件配置
                                case "配置-文件配置":
                                    switch (listObjectHelper.Source.Name)
                                    {
                                        case "视图_库存管理_站点物流周期":
                                            databaseTableName = "`库存管理_站点物流周期`";
                                            id = listObjectHelper.GetValue(index, "id", false);
                                            listObjectHelper.GetValue(index, "站点");
                                            listObjectHelper.GetValue(index, "货代时效(天)");
                                            break;
                                    }
                                    break;
                                #endregion
                                #region 产品-产品类别
                                case "产品-产品类别":
                                    databaseTableName = "`库存管理_产品类别`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "创建负责人", SetValue: VariableHelper.PrincipalName);
                                    listObjectHelper.GetValue(index, "产品名称");
                                    listObjectHelper.GetValue(index, "一级分类");
                                    listObjectHelper.GetValue(index, "二级分类");
                                    listObjectHelper.GetValue(index, "供应商");
                                    listObjectHelper.GetValue(index, "供应商联系方式");
                                    break;
                                #endregion
                                #region 产品-产品信息
                                case "产品-产品信息":
                                    databaseTableName = "`库存管理_产品信息`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "创建负责人", SetValue: VariableHelper.PrincipalName);
                                    listObjectHelper.GetValue(index, "产品名称");
                                    listObjectHelper.GetValue(index, "产品型号");
                                    listObjectHelper.GetValue(index, "产品颜色");
                                    listObjectHelper.GetValue(index, "最长边(CM)");
                                    listObjectHelper.GetValue(index, "次长边(CM)");
                                    listObjectHelper.GetValue(index, "最短边(CM)");
                                    listObjectHelper.GetValue(index, "重量(G)");
                                    listObjectHelper.GetValue(index, "供货周期(天)");
                                    listObjectHelper.GetValue(index, "产品成本(RMB)");
                                    listObjectHelper.GetValue(index, "包装费用(RMB)");
                                    break;
                                #endregion
                                #region 仓库-下单入库
                                case "仓库-下单入库":
                                    databaseTableName = "`库存管理_采购下单`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "下单负责人", SetValue: VariableHelper.PrincipalName);
                                    listObjectHelper.GetValue(index, "产品编号");
                                    listObjectHelper.GetValue(index, "下单数量");
                                    listObjectHelper.GetValue(index, "预计到达日期", SetDate: true);
                                    listObjectHelper.GetValue(index, "入库仓库", SetDefault: VariableHelper.DefaultWarehouse);
                                    listObjectHelper.GetValue(index, "备注");
                                    break;
                                #endregion
                                #region 仓库-入库记录
                                case "仓库-入库记录":
                                    databaseTableName = "`库存管理_仓库入库`";
                                    id = listObjectHelper.GetValue(index, "id", SetFields: "下单编号");
                                    listObjectHelper.GetValue(index, "入库仓库");
                                    listObjectHelper.GetValue(index, "产品编号");
                                    listObjectHelper.GetValue(index, "入库数量");
                                    listObjectHelper.GetValue(index, "黄单编号");
                                    break;
                                #endregion
                                #region 仓库-出库计划
                                case "仓库-出库计划":
                                    databaseTableName = "`库存管理_运营发货`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "发货负责人", SetValue: VariableHelper.PrincipalName);
                                    listObjectHelper.GetValue(index, "SKU");
                                    listObjectHelper.GetValue(index, "发货数量");
                                    listObjectHelper.GetValue(index, "预计到达日期", SetDate: true);
                                    listObjectHelper.GetValue(index, "出库仓库", SetDefault: VariableHelper.DefaultWarehouse);
                                    listObjectHelper.GetValue(index, "发货备注");
                                    break;
                                #endregion
                                #region 仓库-出库记录
                                case "仓库-出库记录":
                                    databaseTableName = "`库存管理_仓库出库`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "SKU");
                                    listObjectHelper.GetValue(index, "出库仓库");
                                    listObjectHelper.GetValue(index, "不良品");
                                    listObjectHelper.GetValue(index, "出库数量");
                                    listObjectHelper.GetValue(index, "货件编号");
                                    listObjectHelper.GetValue(index, "出库备注");
                                    break;
                                #endregion
                                #region 亚马逊-产品负责人
                                case "亚马逊-产品负责人":
                                    databaseTableName = "`库存管理_产品负责人`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "产品负责人", SetValue: VariableHelper.PrincipalName);
                                    listObjectHelper.GetValue(index, "ASIN");
                                    listObjectHelper.GetValue(index, "产品编号");
                                    break;
                                #endregion
                                #region 亚马逊-库龄
                                case "亚马逊-库龄":
                                    databaseTableName = "`库存管理_推广状态`";
                                    id = listObjectHelper.GetValue(index, "id", false);
                                    listObjectHelper.GetValue(index, "SKU");
                                    listObjectHelper.GetValue(index, "推广状态");
                                break;
                                    #endregion
                            }
                            break;
                        case "入库":
                            switch (activeSheet)
                            {
                                #region 仓库-下单入库
                                case "仓库-下单入库":
                                    databaseTableName = "`库存管理_仓库入库`";
                                    id = listObjectHelper.GetValue(index, "id", SetFields: "下单编号");
                                    listObjectHelper.GetValue(index, "产品编号");
                                    listObjectHelper.GetValue(index, "入库数量");
                                    listObjectHelper.GetValue(index, "黄单编号");
                                    listObjectHelper.GetValue(index, "入库仓库", SetDefault: VariableHelper.DefaultWarehouse);
                                    break;
                                #endregion
                            }
                            break;
                        case "出库":
                            switch (activeSheet)
                            {
                                #region 仓库-出库计划
                                case "仓库-出库计划":
                                    databaseTableName = "`库存管理_仓库出库`";
                                    id = listObjectHelper.GetValue(index, "id", SetFields: "发货编号");
                                    listObjectHelper.GetValue(index, "SKU");
                                    listObjectHelper.GetValue(index, "出库仓库");
                                    listObjectHelper.GetValue(index, "不良品");
                                    listObjectHelper.GetValue(index, "出库数量");
                                    listObjectHelper.GetValue(index, "货件编号");
                                    listObjectHelper.GetValue(index, "出库备注");
                                    break;
                                #endregion
                            }
                            break;
                    }
                    if (databaseTableName != "")
                    {
                        string operationType = listObjectHelper.GetValue(index, "操作类型", false);
                        string databaseUpdateField = listObjectHelper.GetSqlUpdateValues();
                        string databaseInsertField = listObjectHelper.GetSqlInsertFields();
                        string databaseInsertData = listObjectHelper.GetSqlInsertValues();
                        switch (operationType)
                        {
                            case "入库":
                            case "出库":
                            case "新增":
                                VariableHelper.MySql.ExecuteNonQuery($"INSERT INTO {databaseTableName}({databaseInsertField}) VALUES ({databaseInsertData})"); break;
                            case "修改":
                                VariableHelper.MySql.ExecuteNonQuery($"UPDATE {databaseTableName} SET {databaseUpdateField} WHERE `id` = {id}"); break;
                            case "删除":
                                VariableHelper.MySql.ExecuteNonQuery($"DELETE FROM {databaseTableName} WHERE `id` = {id}"); break;
                        }
                    }
                }
                if (listObjectHelper.Source.SourceType == XlListObjectSourceType.xlSrcQuery) listObjectHelper.Source.Refresh();
            }
            VariableHelper.Application.ScreenUpdating = true;
            VariableHelper.Application.StatusBar = $"数据更新完成";
        }
        public void RefreshMenu_Click(object sender = null, RibbonControlEventArgs e = null)
        {
            VariableHelper.Application.ActiveWorkbook.RefreshAll();
            ContextMenuHelper.RefreshMenu();
            VariableHelper.Application.ActiveWorkbook.Save();
        }
        public void RibbonMain_Load(object sender = null, RibbonUIEventArgs e = null)
        {
            foreach (string sPrint in PrinterSettings.InstalledPrinters)
            {
                RibbonDropDownItem downItem = Factory.CreateRibbonDropDownItem();
                downItem.Tag = sPrint;
                downItem.Label = sPrint;
                BarcodePrinterName.Items.Add(downItem);
            }
        }
        public void RibbonEditBox_TextChanged(object sender = null, RibbonControlEventArgs e = null)
        {
            ConfigHelper.Set(((RibbonEditBox)sender).Name, ((RibbonEditBox)sender).Text);
            CommonHelper.RibbonMainLoad();
        }
    }
}
