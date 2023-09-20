using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                IRow dataRow = sheet.CreateRow(i+1);
                var cell0 = dataRow.CreateCell(0);
                var cell1 = dataRow.CreateCell(1);
                cell0.SetCellValue(lists[i]);
                cell1.SetCellValue(data[lists[i]].ToString());
            }

            using (FileStream file=new FileStream("C://new.xls",FileMode.Create))
            {
                workBook.Write(file);
            } 

        }
    }
}
