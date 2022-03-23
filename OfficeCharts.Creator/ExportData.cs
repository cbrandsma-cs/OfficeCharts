using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeCharts.Creator
{
    public static class DataSetCreator
    {
        public static IList<ExportData> GenerateData()
        {
            var list = new List<ExportData>();

            for (int i = 0; i < 50; i++)
            {
                var d = new ExportData
                {
                    Id = Guid.NewGuid(),
                    Name = $"Name {i}",
                    Description = $"description field to show off longer text {i}",
                    Value = i % 10
                };

                list.Add(d);
            }

            return list;
        }
    }


    public class ExportData
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int Value { get; set; }

    }
}
