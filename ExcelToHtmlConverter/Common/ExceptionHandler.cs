using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Common
{
    public class ExceptionHandler
    {
        private const string LINE_SEPEARTOR = "====================================================";

        public string WriteToLog(Exception e)
        {
            var builder = new StringBuilder();
            builder
                .AppendLine("An Error occurred in ExcelToHTMLConverter!")
                .AppendLine("TIM: " + DateTime.Now)
                .AppendLine("USR: " + Environment.UserName)
                .AppendLine("PWD: " + Environment.CurrentDirectory)
                .AppendLine();

            if (e != null)
            {
                if (!string.IsNullOrEmpty(e.Message) && !string.IsNullOrEmpty(e.StackTrace))
                {
                    builder
                        .AppendLine(e.Message)
                        .AppendLine(LINE_SEPEARTOR)
                        .AppendLine("StackTrace:")
                        .AppendLine(e.StackTrace)
                        .AppendLine(LINE_SEPEARTOR);
                }
                Exception inner = e.InnerException;
                while (inner != null)
                {
                    builder
                        .AppendLine()
                        .AppendLine("Inner Exception:")
                        .AppendLine();
                    if (!string.IsNullOrEmpty(inner.Message) && !string.IsNullOrEmpty(inner.StackTrace))
                    {
                        builder
                            .AppendLine(inner.Message)
                            .AppendLine(LINE_SEPEARTOR)
                            .AppendLine("StackTrace of Inner Exception:")
                            .AppendLine(inner.StackTrace)
                            .AppendLine(LINE_SEPEARTOR);
                    }
                    inner = inner.InnerException;
                }
            }
            return WriteToLog(builder.ToString());
        }

        public string WriteToLog(string s)
        {
            var filename = GetFilename();
            File.WriteAllText(filename, s);
            return filename;
        }

        public string GetFilename()
        {
            var path = Path.GetTempPath();
            var prefix = "ExcelToHtmlConverter_";
            var timestamp = DateTime.Now
                .ToString("yyyy.MM.dd_HH_mm_ss__ffffzzz")
                .Replace("+", "P")
                .Replace(":", "_");
            var extension = ".txt";
            return path + prefix + timestamp + extension;
        }
    }
}
