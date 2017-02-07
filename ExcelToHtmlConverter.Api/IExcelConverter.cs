using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Api
{
    public interface IExcelConverter
    {
        string ConvertWorksheet(string filename);
    }
}
