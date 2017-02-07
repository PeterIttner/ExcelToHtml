using ExcelToHtmlConverter.Api;
using ExcelToHtmlConverter.Common;
using ExcelToHtmlConverter.Core;
using ExcelToHtmlConverter.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ExcelToHtmlConverter.ViewModels
{
    public class MainViewModel : ModelBase
    {
        #region Memeber

        private IMessageService messageService = DIContainer.Instance.Create<IMessageService>();
        private IExcelConverter excelConverter = DIContainer.Instance.Create<IExcelConverter>();
        private IFileService fileService = DIContainer.Instance.Create<IFileService>();
        private IDialogService dialogService = DIContainer.Instance.Create<IDialogService>();

        #endregion

        #region Ctor

        public MainViewModel()
        {
            this.InfoCommand = new RelayCommand(ExecuteInfo);
            this.CloseCommand = new RelayCommand(ExecuteClose);
            this.ConvertCommand = new RelayCommand(ExecuteConvert, (o) => !string.IsNullOrWhiteSpace(this.Filename));
            this.SelectFileCommand = new RelayCommand(ExecuteSelectFile);
        }

        #endregion

        #region Public Interface

        private string filename;
        public string Filename
        {
            get { return filename; }
            set
            {
                filename = value;
                NotifyPropertyChanged("Filename");
            }
        }

        public ICommand InfoCommand { get; set; }

        public ICommand CloseCommand { get; set; }

        public ICommand ConvertCommand { get; set; }

        public ICommand SelectFileCommand { get; set; }

        #endregion

        #region Command Handler


        private void ExecuteInfo(object obj)
        {
            messageService.ShowInformation("About", buildAboutText());
        }

        private void ExecuteConvert(object obj)
        {
            var html = excelConverter.ConvertWorksheet(Filename);
            var outputFile = string.Format("{0}.html", Filename);
            fileService.WriteToFile(outputFile, html);
            messageService.ShowInformation("Success", "Created " + outputFile);
        }

        private void ExecuteClose(object obj)
        {
            Application.Current.Shutdown();
        }

        private void ExecuteSelectFile(object obj)
        {
            string path;
            if (dialogService.GetFilename("Please provide the file to convert", out path, "Modern Excel files (*.xlsx)|*.xlsx|Old Excel files (*.xls)|*.xls"))
            {
                Filename = path;
            }
        }

        #endregion

        #region Private Helper

        private string buildAboutText()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Excel to HTML Converter");
            sb.AppendLine("Converts Worksheets to HTML files");
            sb.AppendFormat("Version: {0}", Assembly.GetExecutingAssembly().GetName().Version).AppendLine();
            sb.AppendLine("Created by Peter Ittner - ittner.it");
            return sb.ToString();
        }

        #endregion
    }
}
