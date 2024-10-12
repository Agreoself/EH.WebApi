using Microsoft.AspNetCore.Http;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Commons
{
    public class OfficeHelper
    {
        public static void writeToExcel(Dictionary<string, string> data)
        {
            List<string> lists = data.Keys.ToList();
            var workBook = new HSSFWorkbook();
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Company";
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Subject";//主题
            si.Author = "Author";//作者
            workBook.DocumentSummaryInformation = dsi;
            workBook.SummaryInformation = si;

            var sheet = workBook.CreateSheet("Sheet1");

            IRow row = sheet.CreateRow(0);
            var nameCell = row.CreateCell(0);
            var pwdCell = row.CreateCell(1);
            nameCell.SetCellValue("UserName");
            pwdCell.SetCellValue("Password");

            for (int i = 0; i < data.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                var cell0 = dataRow.CreateCell(0);
                var cell1 = dataRow.CreateCell(1);
                cell0.SetCellValue(lists[i]);
                cell1.SetCellValue(data[lists[i]].ToString());
            }

            using FileStream file = new("C://new.xls", FileMode.Create);
            workBook.Write(file);

        }

        public static List<T> ImportFromExcel<T>(IFormFile file) where T : class, new()
        {
            try
            {
                using var stream = file.OpenReadStream();
                var workbook = new XSSFWorkbook(stream); // XSSFWorkbook for .xlsx format
                var sheet = workbook.GetSheetAt(0); // Assuming there is only one sheet

                var data = new List<T>();
                var headerRow = sheet.GetRow(0);
                var cellCount = headerRow.LastCellNum;

                var properties = typeof(T).GetProperties();

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;

                    var obj = new T();

                    for (int j = 0; j < cellCount; j++)
                    {
                        var cell = row.GetCell(j);
                        var cellValue = cell?.ToString() ?? "";

                        var propertyInfo = properties.FirstOrDefault(p => p.Name == headerRow.GetCell(j)?.ToString());
                        if (propertyInfo != null)
                        {
                            var valueType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
                            var safeValue = (cellValue == null) ? null : Convert.ChangeType(cellValue, valueType);
                            propertyInfo.SetValue(obj, safeValue, null);
                        }
                    }

                    data.Add(obj);
                }
                return data;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static byte[] ExportToExcel<T>(List<T> data, string sheetName = "Sheet1") where T : class
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(sheetName);

            // Get properties only from T (excluding base class properties)
            var properties = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public)
                                      .Where(p => p.DeclaringType == typeof(T))
                                      .ToArray();

            // Header row
            var headerRow = sheet.CreateRow(0);
            for (int i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                headerRow.CreateCell(i).SetCellValue(property.Name);
            }

            // Data rows
            for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
            {
                var rowData = data[rowIndex];
                var dataRow = sheet.CreateRow(rowIndex + 1);

                for (int cellIndex = 0; cellIndex < properties.Length; cellIndex++)
                {
                    var property = properties[cellIndex];
                    var value = property.GetValue(rowData);
                    dataRow.CreateCell(cellIndex).SetCellValue(Convert.ToString(value));
                }
            }

            // Convert workbook to byte array
            using var memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
        }
    }
}
