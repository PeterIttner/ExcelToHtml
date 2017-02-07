using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    public class Workbook
    {
        public string Name { get; set; }

        public List<Worksheet> Worksheets { get; set; }
    }
}
