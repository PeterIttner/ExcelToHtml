using System;
using System.Linq;
using System.Collections.Generic;

namespace ExcelToHtmlConverter.Core.Model
{
    /// <summary>
    /// The Worksheet of which a Workbook consists.
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// Creates a new instance of the <see cref="Worksheet"/> class.
        /// Initializes the Id with a unique string.
        /// </summary>
        public Worksheet()
        {
            Id = string.Concat(Guid.NewGuid().ToString("N").Select(c => (char)(c + 17)));
        }

        /// <summary>
        /// Gets or sets the name of the worksheet
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the rows of the Worksheet
        /// </summary>
        public List<Row> Rows { get; set; }

        /// <summary>
        /// Gets or sets the id of the worksheet
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the width of the worksheet in pixel
        /// </summary>
        public int WidthPixel { get; set; }
    }
}