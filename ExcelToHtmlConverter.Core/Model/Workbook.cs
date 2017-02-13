using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    /// <summary>
    /// Excel Workbook
    /// </summary>
    public class Workbook
    {
        /// <summary>
        /// Gets or sets the name of the Workbook
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the Worksheets of the workbook
        /// </summary>
        public List<Worksheet> Worksheets { get; set; }
    }
}
