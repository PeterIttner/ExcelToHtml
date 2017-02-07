using System;
using System.Linq;
using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    public class Worksheet
    {
        public Worksheet()
        {
            Id = string.Concat(Guid.NewGuid().ToString("N").Select(c => (char)(c + 17)));
        }

        public string Name { get; set; }

        public List<Row> Rows { get; set; }

        public string Id { get; set; }

        public int WidthPixel { get; set; }
    }
}