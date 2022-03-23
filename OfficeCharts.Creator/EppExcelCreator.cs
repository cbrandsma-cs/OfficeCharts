using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace OfficeCharts.Creator
{
    public class EppExcelCreator
    {
        public void Create(string filename, IList<ExportData> data)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            using (var package = new ExcelPackage(filename))
            {
                var sheet = package.Workbook.Worksheets.Add("Main");

                var rowIdx = 1;

                sheet.Cells[rowIdx, 1].Value = "ID";
                sheet.Cells[rowIdx, 2].Value = "Name";
                sheet.Cells[rowIdx, 3].Value = "Description";
                sheet.Cells[rowIdx, 4].Value = "Value";

                rowIdx = 2;
                foreach (var d in data)
                {
                    sheet.Cells[rowIdx, 1].Value = d.Id;
                    sheet.Cells[rowIdx, 2].Value = d.Name;
                    sheet.Cells[rowIdx, 3].Value = d.Description;
                    sheet.Cells[rowIdx, 4].Value = d.Value;
                
                    rowIdx++;
                }

                package.Save();
            }

            // using (var workbook = new XLWorkbook())
            // {
            //     var worksheet = workbook.Worksheets.Add("Main");
            //
            //
            //     var row = worksheet.Row(1);
            //
            //     row.Cell(1).Value = "ID";
            //     row.Cell(2).Value = "Name";
            //     row.Cell(3).Value = "Description";
            //     row.Cell(4).Value = "Value";
            //
            //
            //     var rIdx = 2;
            //     foreach (var d in data)
            //     {
            //         row = worksheet.Row(rIdx);
            //         row.Cell(1).Value = d.Id;
            //         row.Cell(2).Value = d.Name;
            //         row.Cell(3).Value = d.Description;
            //         row.Cell(4).Value = d.Value;
            //
            //
            //         rIdx++;
            //     }
            //
            //
            //     // no chart support
            //
            //
            //     workbook.SaveAs(filename);
            // }



        }

    }
}
