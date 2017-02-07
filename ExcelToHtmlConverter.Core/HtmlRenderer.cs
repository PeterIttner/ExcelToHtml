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
        string RenderWorkbook(Workbook workbook);
    }

    public class HtmlRenderer : IRenderer
    {
        private const string TEMPLATE = "Template\\template.html";

        public string RenderWorkbook(Workbook workbook)
        {
            string template = File.ReadAllText(TEMPLATE);
            string html = Engine.Razor.RunCompile(template, "key", typeof(Workbook), workbook);
            return html;
        }
    }
}
