using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;


namespace OfficeCharts.Creator
{
    public class ExcelChartCreator
    {
        public void Create(string filename, IList<ExportData> data)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            var xls = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);


        }
    }


}
