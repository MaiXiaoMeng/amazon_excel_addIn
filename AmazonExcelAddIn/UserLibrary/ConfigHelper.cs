using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AmazonExcelAddIn.UserLibrary
{
    public static class ConfigHelper
    {
        /// <summary>
        /// 读取自定义字符串
        /// </summary>
        /// <param name="Key">名称</param>
        /// <returns>读取不到时返回 Null </returns>
        public static string Get(string Key, string Default = "")
        {
            DocumentProperties properties = (DocumentProperties)VariableHelper.Application.ActiveWorkbook.CustomDocumentProperties;
            foreach (DocumentProperty prop in properties)
            {
                if (prop.Name == Key) { return prop.Value.ToString(); }
            }
            return Default;
        }

        /// <summary>
        /// 写入自定义字符串
        /// </summary>
        /// <param name="Key">名称</param>
        /// <param name="Value">对应值</param>
        public static void Set(string Key, string Value)
        {
            DocumentProperties properties = (DocumentProperties)VariableHelper.Application.ActiveWorkbook.CustomDocumentProperties;
            if (Get(Key) != "")
            {
                properties[Key].Delete();
            }
            properties.Add(Key, false, MsoDocProperties.msoPropertyTypeString, Value);
        }
    }
}
