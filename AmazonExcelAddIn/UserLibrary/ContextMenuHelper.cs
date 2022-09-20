using AmazonExcelAddIn.UserForm;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AmazonExcelAddIn.UserLibrary
{
    public static class ContextMenuHelper
    {
        private static CommandBarButton buttonMenu;
        public static void AddMenu(string menuName, string keyValue, bool cell = false)
        {
            CommandBar contextMenu;
            DelMenu(menuName);
            ListObjectHelper listObjectHelper = new ListObjectHelper("产品-产品菜单", menuName);
            if (cell)
            {
                contextMenu = VariableHelper.Application.CommandBars["cell"];
                foreach (CommandBarControl commandBarControl in contextMenu.Controls)
                {
                    if (commandBarControl.Tag.Length == 32)
                    {
                        commandBarControl.Delete();
                    }
                }
            }
            else
            {
                contextMenu = VariableHelper.Application.CommandBars.Add(menuName, MsoBarPosition.msoBarPopup);
            }

            for (int index = 0; index < listObjectHelper.Source.ListRows.Count; index++)
            {
                string tagContxt = "";
                CommandBarPopup tempContextMenu = null;
                foreach (ListColumn listColumn in listObjectHelper.Source.ListColumns)
                {
                    tagContxt += listObjectHelper.GetValue(index, listColumn.Name);
                    CommandBarPopup popupMenu = (CommandBarPopup)contextMenu.FindControl(MsoControlType.msoControlPopup, Type.Missing, CommonHelper.MD5(tagContxt), true, true);
                    if (popupMenu == null)
                    {
                        if (listColumn.Name != keyValue)
                        {
                            if (tempContextMenu == null)
                            {
                                popupMenu = (CommandBarPopup)contextMenu.Controls.Add(MsoControlType.msoControlPopup, Before: 1);
                            }
                            else
                            {
                                popupMenu = (CommandBarPopup)tempContextMenu.Controls.Add(MsoControlType.msoControlPopup);
                            }
                            popupMenu.Caption = listObjectHelper.GetValue(index, listColumn.Name);
                            popupMenu.Tag = CommonHelper.MD5(tagContxt);
                        }
                        else
                        {
                            if (tempContextMenu == null)
                            {
                                buttonMenu = (CommandBarButton)contextMenu.Controls.Add(MsoControlType.msoControlButton);
                            }
                            else
                            {
                                buttonMenu = (CommandBarButton)tempContextMenu.Controls.Add(MsoControlType.msoControlButton);
                            }
                            buttonMenu.Caption = listObjectHelper.GetValue(index, listColumn.Name);
                            buttonMenu.Tag = "ButtonMenu_Click";
                            buttonMenu.FaceId = 0162;
                            buttonMenu.Click -= ButtonMenu_Click;
                            buttonMenu.Click += ButtonMenu_Click;
                        }
                    }
                    tempContextMenu = popupMenu;
                }
            }
        }
        public static void ButtonMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            int rowNumber;
            string fnSKU;
            ListObjectHelper listObjectHelper;
            switch (Ctrl.Caption)
            {
                #region 更新数据
                case "更新数据":
                    Globals.Ribbons.RibbonMain.RefreshData_Click();
                    break;
                #endregion

                #region 更新菜单
                case "更新菜单":
                    Globals.Ribbons.RibbonMain.RefreshMenu_Click();
                    break;
                #endregion

                #region 随机填充SKU
                case "随机填充SKU":
                    Globals.Ribbons.RibbonMain.RandomSKU_Click();
                    break;
                #endregion

                #region 打印产品条码
                case "打印产品条码":
                    rowNumber = VariableHelper.Application.ActiveCell.Row - 3;
                    listObjectHelper = new ListObjectHelper("仓库-出库计划", "视图_库存管理_出库计划");
                    fnSKU = listObjectHelper.GetValue(rowNumber, "FNSKU");
                    string printQuantity = listObjectHelper.GetValue(rowNumber, "待出库");
                    printQuantity = Interaction.InputBox("请输入需要打印的张数:", DefaultResponse: printQuantity, Title: "提示：");
                    try
                    {
                        DataRow barcodeInformation = VariableHelper.MySql.ExecuteDataRow($"SELECT FNSKU,`短名称` FROM `库存管理_条码信息` WHERE FNSKU = '{fnSKU}'");
                        if (barcodeInformation == null)
                        {
                            new BarcodeCreate(listObjectHelper.GetValue(rowNumber, "FNSKU")).Show();
                        }
                        else
                        {
                            SpireHelper.PrintBarcode(fnSKU, (short)(short.Parse(printQuantity) + 2));
                        }
                    }
                    catch (Exception)
                    {
                        VariableHelper.Application.StatusBar = $"用户取消打印或者打印数量输入错误";
                    }
                    break;
                #endregion

                #region 修改产品条码
                case "修改产品条码":
                    rowNumber = VariableHelper.Application.ActiveCell.Row - 3;
                    listObjectHelper = new ListObjectHelper("仓库-出库计划", "视图_库存管理_出库计划");
                    new BarcodeCreate(listObjectHelper.GetValue(rowNumber, "FNSKU")).Show();
                    break;
                #endregion

                #region 删除产品条码
                case "删除产品条码":
                    fnSKU = VariableHelper.Application.ActiveCell.Value;
                    VariableHelper.MySql.ExecuteNonQuery($"DELETE FROM `库存管理_条码信息` WHERE FNSKU = '{fnSKU}'");
                    break;
                #endregion

                #region 修改头程物流费用
                case "修改头程物流费用":
                    listObjectHelper = new ListObjectHelper("亚马逊-库存", "视图_库存管理_亚马逊库存");
                    Range selectionCellRange = VariableHelper.Application.Selection;
                    foreach (Range range in selectionCellRange.Rows)
                    {
                        rowNumber = range.Row - 3;
                        string SKU = listObjectHelper.GetValue(rowNumber, "SKU");
                        int headCost = Convert.ToInt32(listObjectHelper.GetValue(rowNumber, "头程费用"));
                        DataRow productInformation = VariableHelper.MySql.ExecuteDataRow($"SELECT SKU,`费用(RMB)` FROM `库存管理_头程物流费用` WHERE SKU = '{SKU}'");
                        if (productInformation == null)
                        {
                            VariableHelper.MySql.ExecuteNonQuery($"INSERT INTO `库存管理_头程物流费用`(`SKU`, `费用(RMB)`) VALUES('{SKU}', '{headCost}')");
                        }
                        else
                        {
                            VariableHelper.MySql.ExecuteNonQuery($"UPDATE `库存管理_头程物流费用` SET `费用(RMB)` = '{headCost}' WHERE `SKU` = '{SKU}'");
                        }
                    }
                    listObjectHelper.Source.Refresh();
                    VariableHelper.Application.StatusBar = $"头程物流费用修改完成";
                    break;
                #endregion

                #region 其他标签类型
                default:
                    VariableHelper.Application.ActiveCell.Value = Ctrl.Caption;
                    break;
                    #endregion
            }
        }
        public static void DelMenu(string menuName)
        {
            try { VariableHelper.Application.CommandBars[menuName].Delete(); } catch (Exception) { }
        }
        public static void RefreshMenu()
        {
            VariableHelper.Application.StatusBar = $"正在更新菜单数据";
            AddMenu("视图_库存管理_菜单管理_产品菜单", "id");
            AddMenu("视图_库存管理_菜单管理_类别菜单", "产品名称");
            AddMenu("视图_库存管理_菜单管理_SKU菜单", "SKU");
            AddMenu("视图_库存管理_菜单管理_条码菜单", "条码菜单");
            AddMenu("视图_库存管理_菜单管理_头程菜单", "头程菜单");
            AddMenu("视图_库存管理_菜单管理_右键菜单", "三级菜单",cell: true);
            VariableHelper.Application.StatusBar = $"菜单数据更新完成";
        }
    }
}
