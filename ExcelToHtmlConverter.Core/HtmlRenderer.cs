using ExcelToHtmlConverter.Core.Model;
using RazorEngine;
using RazorEngine.Templating;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Core
{
    public interface IRenderer
    {
        string RenderWorkbook(Workbook workbook, string templateFile);
    }

    public class HtmlRenderer : IRenderer
    {
        public string RenderWorkbook(Workbook workbook, string templateFile)
        {
            string template = File.ReadAllText(templateFile);
            string html = Engine.Razor.RunCompile(template, "key", typeof(Workbook), workbook);
            return html;
        }
    }
}
