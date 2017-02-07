using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Services
{
    public interface IDialogService
    {
        bool GetFilename(string description, out string path, string filter = "");
    }

    public class DialogService : IDialogService
    {
        public bool GetFilename(string description, out string path, string filter = "")
        {
            path = string.Empty;

            var dlg = new OpenFileDialog();
            dlg.CheckFileExists = true;
            dlg.CheckPathExists = true;
            dlg.Filter = filter;
            dlg.DereferenceLinks = false;
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
            dlg.Multiselect = false;
            dlg.Title = description;

            var result = dlg.ShowDialog();
            if (result.HasValue && result.Value)
            {
                path = dlg.FileName;
                return true;
            }
            return false;
        }
    }
}
