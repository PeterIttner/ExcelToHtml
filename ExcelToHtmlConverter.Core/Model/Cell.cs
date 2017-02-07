using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    public class Cell
    {
        public string Value { get; set; }
        public string BackgroundColor { get; set; }
        public string TextColor { get; set; }
        public string WidthClass { get; set; }
        public int WidthPixel { get; set; }
    }
}