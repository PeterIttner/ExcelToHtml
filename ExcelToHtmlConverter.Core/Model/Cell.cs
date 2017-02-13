using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    /// <summary>
    /// A Cell of a Worksheet
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Gets or sets the content of the cell
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Gets or sets the BackgroundColor
        /// </summary>
        public string BackgroundColor { get; set; }

        /// <summary>
        /// Gets or sets the TextColor
        /// </summary>
        public string TextColor { get; set; }

        /// <summary>
        /// Gets or sets the width class (css-class)
        /// </summary>
        public string WidthClass { get; set; }

        /// <summary>
        /// Gets or sets the width in pixel
        /// </summary>
        public int WidthPixel { get; set; }
    }
}