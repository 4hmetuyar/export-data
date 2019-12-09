using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Threading.Tasks;
using System.Web;

namespace GenerateReport
{
    public class ExcelExport<T>
    {
        /// <summary>
        /// Export excel generic type
        /// </summary>
        /// <param name="enumerable"></param>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static async Task<bool> ToExcelFile(IEnumerable<T> enumerable, string fileName, string sheetName)
        {
            try
            {
                Guid newGuid = Guid.NewGuid();

                if (fileName == null) fileName = newGuid.ToString();
                if (sheetName == null) sheetName = "sheet1";
                DataTable convertToDataTable = await ConvertToDataTable(enumerable);
                using (XLWorkbook xlWorkbook = new XLWorkbook())
                {
                    xlWorkbook.Worksheets.Add(convertToDataTable, sheetName);
                    var filePath = HttpContext.Current.Server.MapPath($"/content/upload/{fileName}.xlsx");
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
        public static async Task<DataTable> ConvertToDataTable<T>(IEnumerable<T> data)
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
