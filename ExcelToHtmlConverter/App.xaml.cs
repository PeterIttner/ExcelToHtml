using ExcelToHtmlConverter.Common;
using System;
using System.IO;
using System.Text;
using System.Windows;

namespace ExcelToHtmlConverter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var filename = string.Empty;
            if (e.ExceptionObject is Exception)
            {
                filename = new ExceptionHandler().WriteToLog(e.ExceptionObject as Exception);
            }
            else
            {
                filename = new ExceptionHandler().WriteToLog("An unexpected error occurred without any further information.");
            }
            ShowErrorMessage(filename);
            Environment.Exit(1);
        }

        private void ShowErrorMessage(string filename)
        {
            MessageBox.Show(new StringBuilder()
                .AppendLine("An error occurred.")
                .AppendLine("Please see the following logfile for more information.")
                .AppendLine(filename)
                .AppendLine()
                .AppendLine("The application will now close...")
                .ToString()
                , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
