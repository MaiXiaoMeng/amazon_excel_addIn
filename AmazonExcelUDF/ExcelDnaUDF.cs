using ExcelDna.Integration;
using System.Windows.Forms;
using ExcelPIA = Microsoft.Office.Interop.Excel;


namespace AmazonExcelUDF
{
    /// <summary>
    /// ExcelDNA 自定义函数
    /// </summary>
    public static class ExcelDnaUDF
    {
        //*****************************************************************************
        //让自定义函数在工作表函数中可见：必须有返回类型
        //void类型的函数在工作表函数向导中是不可见的,也不可在VBA中执行。
        //viod类型的函数，如果带有参数的话，在宏表函数向导中可见；
        //非viod函数或Void不带参数的话，在宏表函数向导中不可见。（实际上是命令,Type=2）
        //IsMacroType = true ，ReferenceToRange函数才起作用，否则运行时报错
        //VBA中调用方法：ret = Application.Run("Addthem", 1, 2, Range("A1"))
        //Application.ExecuteExcel4Macro ("xxxx()")        
        //Application.ExecuteExcel4Macro ("UDF(""" & ThisWorkbook.Path & Application.PathSeparator & "mystring" & """)")  //传递字符串变量
        //Application.ExecuteExcel4Macro ("UNREGISTER(""" & "xxxx.XLL" & """)") //传递字符串变量

        //https://msdn.microsoft.com/zh-cn/library/office/bb687837.aspx  C API 函数参考 ，Excel 2013 XLL SDK API 函数引用

        //在XP和Office2003中，VBA的Application.RegisteredFunctions函数、Application.RegisterXLL函数
        //区分大小写，所以要确保加载XLL函数的名称字母大小写与XLL文件的名字一致

        //ExcelFunction属性
        //Name          如果未给出，则使用实际的方法名称
        //Description   如果不使用，说明将不包含在文档中
        //Category      如果未提供功能，则将归入"功能" 
        //HelpTopic     可用于在Excel的功能向导中将功能链接到生成的帮助

        //IsHidden          是否隐藏该函数，设置为true时，意味着在函数管理器上隐藏
        //IsVolatile        是否为易失性函数，设置为true时，意味着该函数会实时刷新
        //IsMacroType       是否使用宏类型,设置为true时，ExcelDNA注册该函数，会调用xlfRegister
        //IsThreadSafe      是否线程安全，设置为true时，意味着你的函数安全的多线程重新计算，如果在注册字符串最后加上"$"符号，可以在内部调用xlfRegister
        //IsClusterSafe     是否集群安全，设置为true时，意味着你的函数在集群时安全
        //IsExceptionSafe   是否异常安全，设置为true时，意味着无论何时出现未知的异常时，Excel应该崩溃，该参数最好忽略

        //ExplicitRegisration 是否为显示注册，设置为true时不会自动注册这个函数

        //ExcelArgument属性
        //Name          如果未给出，则使用实际的参数名称
        //Description   如果不使用，说明将不包含在文档中
        //AllowReference 设置为true时为Range引用,否则为二维数组

        //ExcelCommand属性
        //MenuName      命令菜单
        //MenuText      菜单命令
        //Name          如果未给出，则使用实际的参数名称
        //Description   如果不使用，说明将不包含在文档中
        //HelpTopic     可用于在Excel的功能向导中将功能链接到生成的帮助
        //ShortCut      如果不使用，文档中将不包含快捷方式

        //ExcelFunctionDoc属性
        //Name          如果未给出，则使用实际的方法名称
        //Description   如果不使用，说明将不包含在文档中
        //Category      如果未提供功能，则将归入“ 功能”
        //HelpTopic     可用于在Excel的功能向导中将功能链接到生成的帮助
        //Returns       返回值说明
        //Summary       文档中包含的功能的详细讨论
        //Remarks       使用说明和/或可能的错误
        //Example       示例代码演示正确用法

        //如果包含ExcelDna.Documentation作为参考（NuGet包中的默认值），则可以使用其他属性ExcelFunctionDoc来代替该ExcelFunction属性，该属性包括可以用于其他文档的其他字段。

        //*****************************************************************************

        #region XLL的自定义函数
        //http://yi-lee.blog.163.com/blog/static/4955152620151171395919/
        /// <summary>
        /// Addition UDF function
        /// </summary>
        /// <param name="Param1"></param>
        /// <param name="Param2"></param>
        /// <param name="Param_range"></param>
        /// <returns></returns>
        [ExcelFunction(Description = "Addition function",
                       Category = "ExcelDNA Demo Function",
                       HelpTopic = "http://club.excelhome.net/thread-1025191-1-1.html", //HelpTopic="MyHelp.chm!102"
                       IsHidden = false,
                       IsVolatile = true,
                       IsMacroType = true,
                       Name = "AddThem")]
        public static object AddThem([ExcelArgument(Description = "DecimalValue1 or Range1", Name = "decimal1")] object Param1,
                                     [ExcelArgument(Description = "DecimalValue2 or Range2", Name = "decimal2")] object Param2)
        {
            double value1, value2;

            if (ExcelDnaUtil.IsInFunctionWizard()) return "Waiting for click on wizard ok button to calculate.";

            if (Param1 is ExcelMissing) return ExcelError.ExcelErrorNA;
            if (Param2 is ExcelMissing) return ExcelError.ExcelErrorNA;

            if (Param1 is ExcelEmpty)
            {
                value1 = 0;
            }
            else
            {
                if (!IsNumberic(Param1.ToString(), out value1)) return (object)ExcelError.ExcelErrorValue;
            }
            if (Param2 is ExcelEmpty)
            {
                value2 = 0;
            }
            else
            {
                if (!IsNumberic(Param2.ToString(), out value2)) return (object)ExcelError.ExcelErrorValue;
            }
            return value1 + value2;
        }
        #endregion XLL的自定义函数

        #region excel command 加载项菜单

        /// <summary>
        /// menu command
        /// </summary>
        [ExcelCommand(MenuName = "Demo Tools", MenuText = "Square Selection")]
        public static void SquareRange()
        {
            object[,] result;

            // Get a reference to the current selection
            ExcelReference selection = null; ;
            try
            {
                selection = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            }
            catch
            {
                return;
            }
            if (selection == null) return;

            dynamic obj = ReferenceToRange(selection);
            if (!(obj is ExcelPIA.Range)) return;

            // Get the value of the selection
            object selectionContent = selection.GetValue();
            if (selectionContent is object[,])
            {
                object[,] values = (object[,])selectionContent;
                int rows = values.GetLength(0);
                int cols = values.GetLength(1);
                result = new object[rows, cols];

                // Process the values
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        if (values[i, j] is double)
                        {
                            double val = (double)values[i, j];
                            result[i, j] = val * val;
                        }
                        else
                        {
                            result[i, j] = values[i, j];
                        }
                    }
                }
            }
            else if (selectionContent is double)
            {
                double value = (double)selectionContent;
                result = new object[,] { { value * value } };
            }
            else
            {
                result = new object[,] { { "Selection was not a range or a number, but " + selectionContent.ToString() } };
            }

            // Now create the target reference that will refer to Sheet 1, getting a reference that contains the SheetId first
            ExcelReference sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, "Sheet1"); // Throws exception if no Sheet1 exists
            // ... then creating the reference with the right size as new ExcelReference(RowFirst, RowLast, ColFirst, ColLast, SheetId)
            int resultRows = result.GetLength(0);
            int resultCols = result.GetLength(1);
            ExcelReference target = new ExcelReference(0, resultRows - 1, 0, resultCols - 1, sheet.SheetId);
            // Finally setting the result into the target range.
            target.SetValue(result);
        }

        /// <summary>
        /// menu command
        /// </summary>
        [ExcelCommand(MenuName = "Demo Tools", MenuText = "Sum Selection")]
        public static void SumRange()
        {
            //xlcall.h
            //https://code.msdn.microsoft.com/Excel-2010-Writing-791e9222/sourcecode?fileId=25565&pathId=1844590411

            // Get a reference to the current selection
            ExcelReference selection = null; ;
            try
            {
                selection = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            }
            catch
            {
                return;
            }
            if (selection == null) return;

            object sum = XlCall.Excel(XlCall.xlfSum, selection);

            ExcelReference target = new ExcelReference(0, 0, 0, 0);  //activesheet.cells[1,1]
            target.SetValue(sum);
        }

        /// <summary>
        /// test RunTagMacro
        /// </summary>
        [ExcelCommand(MenuName = "Demo Tools", MenuText = "RunTagMacroTest", Name = "CmdName")]
        public static void RunTagTest()
        {
            //这个函数测试了ExcelDna的RunTagMacro功能，此功能通过Application.Run调用control.Tag命名的vba函数或command
            MessageBox.Show("Test RunTagMacro command!");
        }

        #endregion excel command 加载项菜单

        private static bool IsNumberic(string str, out double vsNum)
        {
            bool isNum;
            isNum = double.TryParse(str, System.Globalization.NumberStyles.Float,
                System.Globalization.NumberFormatInfo.InvariantInfo, out vsNum);
            return isNum;
        }

        /// <summary>
        /// ExcelReference to Range
        /// </summary>
        /// <param name="xlref"></param>
        /// <returns></returns>
        private static dynamic ReferenceToRange(ExcelReference xlref)          //简版的
        {
            dynamic app = ExcelDnaUtil.Application;
            return app.Range[XlCall.Excel(XlCall.xlfReftext, xlref, true)];
        }

        [ExcelFunction(Description = "获取单元格地址", Category = "MaiXiaoMeng",
            IsHidden = false,
            IsMacroType = true,
            IsVolatile = false,
            IsThreadSafe = false,
            Name = "RangeAddress"
            )
        ]
        public static object CallerAddress()
        {
            ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            string refText = (string)XlCall.Excel(XlCall.xlfReftext, caller, true);
            dynamic app = ExcelDnaUtil.Application;
            dynamic Range = app.Range[refText];
            return Range.Address;
        }

        [ExcelFunction(Name = "XlfVlookup", Category = "MaiXiaoMeng", IsMacroType = false, IsHidden = false, IsVolatile = false)]
        public static object XlfVlookup(
         [ExcelArgument(Name = "Lookup_value")] object Lookup_value,
         [ExcelArgument(Name = "Table_array", AllowReference = true)] object Table_array,
         [ExcelArgument(Name = "Col_index_num")] object Col_index_num,
         [ExcelArgument(Name = "Range_lookup")] object Range_lookup
        )
        {
            return XlCall.Excel(XlCall.xlfVlookup, Lookup_value, Table_array, Col_index_num, Range_lookup);
        }

        [ExcelFunction(Description = "可变参数求和(Ojbect)", Category = "ExcelDNA Demo Function", ExplicitRegistration = true)]
        public static object ParamsSUM(params object[] values)
        {
            return "Ojbect" + values.Length.ToString();
        }
    }
}
