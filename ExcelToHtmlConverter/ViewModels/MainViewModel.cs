using ExcelToHtmlConverter.Api;
using ExcelToHtmlConverter.Common;
using ExcelToHtmlConverter.Services;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ExcelToHtmlConverter.ViewModels
{
    /// <summary>
    /// MainViewModel that handles the complete business logic of the View.
    /// </summary>
    public class MainViewModel : ModelBase
    {
        #region Memeber

        private IMessageService messageService = DIContainer.Instance.Create<IMessageService>();
        private IExcelConverter excelConverter = DIContainer.Instance.Create<IExcelConverter>();
        private IFileService fileService = DIContainer.Instance.Create<IFileService>();
        private IDialogService dialogService = DIContainer.Instance.Create<IDialogService>();

        private string filename;
        private string template;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates a new instance of the <see cref="MainViewModel"/> class.
        /// </summary>
        public MainViewModel()
        {
            this.InfoCommand = new RelayCommand(ExecuteInfo);
            this.CloseCommand = new RelayCommand(ExecuteClose);
            this.ConvertCommand = new RelayCommand(ExecuteConvert, (o) => !string.IsNullOrWhiteSpace(this.Filename));
            this.SelectFileCommand = new RelayCommand(ExecuteSelectFile);
            this.SelectTemplateCommand = new RelayCommand(ExecuteSelectTemplate);
        }

        #endregion

        #region Public Interface

        /// <summary>
        /// Gets or sets the selected filename of the Excel file that should be converted.
        /// </summary>
        public string Filename
        {
            get { return filename; }
            set
            {
                filename = value;
                NotifyPropertyChanged("Filename");
            }
        }

        /// <summary>
        /// Gets or sets the selected template file that should be used for the rendering.
        /// </summary>
        public string Template
        {
            get { return template; }
            set
            {
                template = value;
                NotifyPropertyChanged("Template");
            }
        }

        /// <summary>
        /// Displays the about information to the user.
        /// </summary>
        public ICommand InfoCommand { get; set; }

        /// <summary>
        /// Closes the whole application
        /// </summary>
        public ICommand CloseCommand { get; set; }

        /// <summary>
        /// Starts the converstion of the selected file.
        /// Can only be called when a file has been selected.
        /// </summary>
        public ICommand ConvertCommand { get; set; }

        /// <summary>
        /// Opens a dialog, so that the user can select the file to convert there.
        /// </summary>
        public ICommand SelectFileCommand { get; set; }

        /// <summary>
        /// Opens a dialog, so that the user can select a custom template there.
        /// </summary>
        public ICommand SelectTemplateCommand { get; set; }

        #endregion

        #region Command Handler

        private void ExecuteInfo(object obj)
        {
            messageService.ShowInformation("About", BuildAboutText());
        }

        private void ExecuteConvert(object obj)
        {
            var html = string.IsNullOrEmpty(Template) ? excelConverter.ConvertWorksheet(Filename) : excelConverter.ConvertWorksheet(Filename, Template);
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

        private void ExecuteSelectTemplate(object obj)
        {
            string path;
            if (dialogService.GetFilename("Please provide the template file", out path))
            {
                Template = path;
            }
        }

        #endregion

        #region Private Helper

        private string BuildAboutText()
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
