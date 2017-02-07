using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelToHtmlConverter.Services
{
    public interface IMessageService
    {
        void ShowInformation(string title, string message);
    }

    public class MessageService : IMessageService
    {
        public void ShowInformation(string title, string message)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
