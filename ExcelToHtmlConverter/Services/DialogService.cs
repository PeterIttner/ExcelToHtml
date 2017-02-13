using Microsoft.Win32;
using System;

namespace ExcelToHtmlConverter.Services
{
    /// <summary>
    /// Interface for dialog abstraction services.
    /// </summary>
    public interface IDialogService
    {
        /// <summary>
        /// Gets the path to a file.
        /// </summary>
        /// <param name="description">The description that should be displayed to the user</param>
        /// <param name="path">The selected path by the user</param>
        /// <param name="filter">(Optional) The file-filter that should used (example: *.xls)</param>
        /// <returns>true, if the user selected a file path, false otherwise</returns>
        bool GetFilename(string description, out string path, string filter = "");
    }

    /// <summary>
    /// Shows Dialogs to the user
    /// </summary>
    public class DialogService : IDialogService
    {
        /// <summary>
        /// Gets the path to a file.
        /// </summary>
        /// <param name="description">The description that should be displayed to the user</param>
        /// <param name="path">The selected path by the user</param>
        /// <param name="filter">(Optional) The file-filter that should used (example: *.xls)</param>
        /// <returns>true, if the user selected a file path, false otherwise</returns>
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
