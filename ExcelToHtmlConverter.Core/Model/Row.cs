using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    /// <summary>
    /// A Row of Cells in a Worksheet
    /// </summary>
    public class Row
    {
        /// <summary>
        /// Gets or sets the cells of the row
        /// </summary>
        public List<Cell> Cells { get; set; }
    }
}