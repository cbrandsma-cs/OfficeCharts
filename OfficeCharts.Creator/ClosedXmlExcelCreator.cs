using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using ClosedXML.Excel;


namespace OfficeCharts.Creator
{
    public class ClosedXmlExcelCreator
    {
        public void Create(string filename, IList<ExportData> data)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Main");

                
                var row = worksheet.Row(1);

                row.Cell(1).Value = "ID";
                row.Cell(2).Value = "Name";
                row.Cell(3).Value = "Description";
                row.Cell(4).Value = "Value";


                var rIdx = 2;
                foreach (var d in data)
                {
                    row = worksheet.Row(rIdx);
                    row.Cell(1).Value = d.Id;
                    row.Cell(2).Value = d.Name;
                    row.Cell(3).Value = d.Description;
                    row.Cell(4).Value = d.Value;


                    rIdx++;
                }


                // no chart support


                workbook.SaveAs(filename);
            }



        }


    }


}
