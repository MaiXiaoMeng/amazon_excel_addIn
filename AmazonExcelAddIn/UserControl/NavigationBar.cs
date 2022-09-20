
﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using AmazonExcelAddIn.UserLibrary;
using System.Diagnostics;

namespace AmazonExcelAddIn.UserControl
{
    [ToolboxItem(false)]
    public partial class NavigationBar : System.Windows.Forms.UserControl
    {
        public NavigationBar()
        {
            InitializeComponent();
        }

        private string GetMenuName(string menuName)
        {
            return new string(' ', 12 - menuName.Length) + menuName;
        }

        private void NavigationBar_Load(object sender, EventArgs e)
        {
            TVMenu.Nodes.Add("首页", GetMenuName("首页"));
            foreach (Excel.Worksheet worksheet in VariableHelper.Application.ActiveWorkbook.Worksheets)
            {
                string[] categoryName = worksheet.Name.Split('-');
                if (categoryName.Length > 1)
                {
                    TreeNode[] menuNode = TVMenu.Nodes.Find(categoryName[0], false);
                    if (menuNode.Length > 0)
                    {
                        menuNode[0].Nodes.Add(categoryName[1], worksheet.Name);
                    }
                    else
                    {
                        TVMenu.Nodes.Add(categoryName[0], GetMenuName(categoryName[0])).Nodes.Add(categoryName[1], worksheet.Name);
                    }
                }
            }
           
        }

        private void NavigationBar_SizeChanged(object sender, EventArgs e)
        {
            TVMenu.Height = ((Control)sender).Height;
        }

        private void TVMenu_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string[] categoryName = e.Node.Text.Trim().Split('-');
            if (categoryName.Length>1|| e.Node.Text.Trim() == "首页")
            {
                VariableHelper.Application.ActiveWorkbook.Sheets[e.Node.Text.Trim()].Activate();
            }
            
        }
    }
}
