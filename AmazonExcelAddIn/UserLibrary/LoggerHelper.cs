using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace AmazonExcelAddIn.UserLibrary
{
    public static class LoggerHelper
    {
        /// <summary>
        /// 严重 - 严重的错误事件，将会导致应用程序的退出，需要运维管理人员马上介入
        /// </summary>
        public static void Fatal(string msg, [CallerMemberName] string memberName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToLocalTime()} {Path.GetFileName(filePath)} > {memberName}({lineNumber}) {msg}", "Fatal");
        }
        /// <summary>
        /// 错误 - 错误事件，影响正常使用，但仍然不影响系统的继续运行
        /// </summary>
        public static void Error(string msg, [CallerMemberName] string memberName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToLocalTime()} {Path.GetFileName(filePath)} > {memberName}({lineNumber}) {msg}", "Error");

        }
        /// <summary>
        /// 警告 - 预期之外的运行状况，可能会出现潜在错误的情形，比如大量时延过大等；一般是由系统资源等技术原因触发
        /// </summary>
        public static void Warn(string msg, [CallerMemberName] string memberName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToLocalTime()} {Path.GetFileName(filePath)} > {memberName}({lineNumber}) {msg}", "Warn");
        }
        /// <summary>
        /// 提示 - 粗粒度记录应用程序的正常运行过程中的关键信息
        /// </summary>
        public static void Info(string msg, [CallerMemberName] string memberName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToLocalTime()} {Path.GetFileName(filePath)} > {memberName}({lineNumber}) {msg}", "Info");
        }
        /// <summary>
        /// 调试 - 细粒度记录应用程序的正常运行过程中的信息，帮助调试和诊断应用程序
        /// </summary>
        public static void Debug(string msg, [CallerMemberName] string memberName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToLocalTime()} {Path.GetFileName(filePath)} > {memberName}({lineNumber}) {msg}", "Debug");
        }

        /// <summary>
        /// 跟踪 - 细粒度记录应用程序的正常运行过程中的信息，帮助调试和诊断应用程序
        /// </summary>
        public static void Trace(string msg, [CallerMemberName] string memberName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0)
        {
            System.Diagnostics.Debug.WriteLine($"{DateTime.Now.ToLocalTime()} {Path.GetFileName(filePath)} > {memberName}({lineNumber}) {msg}", "Trace");
        }
    }
}
