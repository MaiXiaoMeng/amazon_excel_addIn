using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AmazonExcelAddIn.UserLibrary
{
    public class ListObjectHelper
    {
        public Excel.ListObject Source { get; set; }
        public string SqlUpdateValue { get; set; }
        public string SqlInsertFields { get; set; }
        public string SqlInsertValues { get; set; }
        public ListObjectHelper(string sheetName,string listObjectsName)
        {
            Source = VariableHelper.Application.Worksheets[sheetName].ListObjects[listObjectsName];
        }

        public string GetValue(int RowInd, string ColName, bool SetSql = true, string SetFields = "", string SetValue = "", string SetDefault = "", bool SetDate = false)
        {
            string Value;
            try
            {
                int ColInd = Source.ListColumns.get_Item(ColName).Index;
                Value = ((Excel.Range)Source.Range.get_Item(RowInd + 2, ColInd)).Value2.ToString();
                if (SetDate)
                {
                    Value = DateTime.FromOADate(double.Parse(Value)).ToString("yyyy-MM-dd HH:mm:ss");

                }
                if (Value == "")
                {
                    Value = SetDefault;
                }
            }
            catch (Exception)
            {
                Value = SetDefault;
            }

            if (SetSql)
            {
                string _ColName = ColName;
                string _Value = Value;
                if (SetFields.Length > 0)
                {
                    _ColName = SetFields;

                }
                if (SetValue.Length > 0)
                {
                    _Value = SetValue;
                }
                if (_Value.Length > 0)
                {
                    SqlUpdateValue = $"{SqlUpdateValue} `{_ColName}` = '{_Value}',";
                    SqlInsertFields = $"{SqlInsertFields} `{_ColName}`,";
                    SqlInsertValues = $"{SqlInsertValues} '{_Value}',";
                }
            }
            return Value;

        }

        public int GetRowsLength()
        {
            return Source.ListRows.Count;
        }

        public string GetSqlUpdateValues()
        {
            string SqlValue = SqlUpdateValue.Substring(0, SqlUpdateValue.Length - 1);
            SqlUpdateValue = "";
            return SqlValue;
        }
        public string GetSqlInsertFields()
        {
            string _SqlInsertFields = SqlInsertFields.Substring(0, SqlInsertFields.Length - 1);
            SqlInsertFields = "";
            return _SqlInsertFields;
        }

        public string GetSqlInsertValues()
        {
            string _SqlInsertValues = SqlInsertValues.Substring(0, SqlInsertValues.Length - 1);
            SqlInsertValues = "";
            return _SqlInsertValues;
        }

    }
}
