using ExcelToHtmlConverter.Core.Model;
using RazorEngine;
using RazorEngine.Templating;
using System.IO;

namespace ExcelToHtmlConverter.Core
{
    /// <summary>
    /// Redering abstraction for rendering a Workbook.
    /// </summary>
    public interface IRenderer
    {
        /// <summary>
        /// Renders the given workbook with template file.
        /// </summary>
        /// <param name="workbook">The workbook that should be rendered</param>
        /// <param name="templateFile">The template that should be used</param>
        /// <returns>The rendered workbook as string</returns>
        string RenderWorkbook(Workbook workbook, string templateFile);
    }

    /// <summary>
    /// Renderer for rendering Workbooks to HTML
    /// </summary>
    public class HtmlRenderer : IRenderer
    {
        /// <summary>
        /// Renders the given workbook with template file.
        /// The rendering engine is RazorTemplate engine.
        /// </summary>
        /// <param name="workbook">The workbook that should be rendered</param>
        /// <param name="templateFile">The template that should be used</param>
        /// <returns>The rendered workbook as string</returns>
        public string RenderWorkbook(Workbook workbook, string templateFile)
        {
            string template = File.ReadAllText(templateFile);
            string html = Engine.Razor.RunCompile(template, "key", typeof(Workbook), workbook);
            return html;
        }
    }
}
