
namespace ExcelToHtmlConverter.Api
{
    /// <summary>
    /// Abstraction for converting excel sheets
    /// </summary>
    public interface IExcelConverter
    {
        /// <summary>
        /// Converts the worksheet from the given file to a string.
        /// </summary>
        /// <param name="filename">The path to the worksheet file</param>
        /// <returns>The converted worksheet as string</returns>
        string ConvertWorksheet(string filename);

        /// <summary>
        /// Converts the worksheet from the given file to a string.
        /// Applies the given template file for the conversion.
        /// </summary>
        /// <param name="filename">The path to the worksheet file</param>
        /// <param name="templateFile">The path to the template file that is used for the conversion</param>
        /// <returns>The converted worksheet as string</returns>
        string ConvertWorksheet(string filename, string templateFile);
    }
}
