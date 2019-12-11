using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace GenerateReport
{
    public static class ExcelExport<T>
    {
        /// <summary>
        /// Export excel generic type
        /// </summary>
        /// <param name="enumerable"></param>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool ToExcelFile(IEnumerable<T> enumerable, string filePath, string sheetName)
        {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            try
            {
                if (sheetName == null) sheetName = "sheet1";
                DataTable convertToDataTable = ConvertToDataTable(enumerable);
                using (XLWorkbook xlWorkbook = new XLWorkbook())
                {
                    xlWorkbook.Worksheets.Add(convertToDataTable, sheetName);
                    xlWorkbook.SaveAs(filePath);
                }
                return true;
            }
            catch (Exception e)
            {
                //Logging.
                return false;
            }
        }

        /// <summary>
        /// Convert Data Table for Generic Type
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        private static DataTable ConvertToDataTable<T>(IEnumerable<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }
    }
}
