using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelToHtmlConverter.Core;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.IO;
using System.Diagnostics;
using System.Text;

namespace ExcelToHtmlConverterTest.Core
{
    [TestClass]
    public class ExcelConverterTest
    {
        private const string TESTFILE = "Assets\\TestSheet1.xlsx";

        [TestMethod]
        public void TestThat_ConvertWorksheetToHtml_WorksFine()
        {
            var converter = new ExcelConverter();
            var content = converter.ConvertWorksheet(TESTFILE);
            File.WriteAllText(TESTFILE + ".html", content, Encoding.UTF8);
            Process.Start(TESTFILE + ".html");
        }
    }
}
