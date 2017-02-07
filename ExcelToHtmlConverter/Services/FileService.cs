using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Services
{
    public interface IFileService
    {
        void WriteToFile(string filename, string content);
    }

    public class FileService : IFileService
    {
        public void WriteToFile(string filename, string content)
        {
            File.WriteAllText(filename, content);
        }
    }
}
