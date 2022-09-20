using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Linq;

namespace AmazonExcelUDF
{
    public class AddIn : IExcelAddIn
    {
        public void AutoClose()
        {
            XlCall.Excel(XlCall.xlcAlert, "AutoClose");
        }

        public void AutoOpen()
        {
            // 注册自定义函数的智能互交
            ExcelComAddInHelper.LoadComAddIn(new ComAddInConnection());

            // 注册自定义支持可变参数函数
            ExcelRegistration.GetExcelFunctions()
                .Where(func => func.FunctionAttribute.Name.StartsWith("ParamsSUM"))
                .ProcessParamsRegistrations()
                .RegisterFunctions();

        }
    }
}
