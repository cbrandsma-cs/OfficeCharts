using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeCharts.Creator;

namespace OfficeCharts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var data = DataSetCreator.GenerateData();

            var creator = new ClosedXmlExcelCreator();
            creator.Create("c:\\git\\ExcelClosedXml.xlsx", data);

            var creator2 = new EppExcelCreator();
            creator2.Create("c:\\git\\ExcelEPPlusExample2.xlsx", data);

        }
    }
}
