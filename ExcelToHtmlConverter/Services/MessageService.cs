using System.Windows;

namespace ExcelToHtmlConverter.Services
{
    /// <summary>
    /// Interface for abstracting user notifications
    /// </summary>
    public interface IMessageService
    {
        /// <summary>
        /// Shows a message with the given title and text.
        /// </summary>
        /// <param name="title">The title of the message</param>
        /// <param name="message">The text of the message</param>
        void ShowInformation(string title, string message);
    }

    /// <summary>
    /// Handles notifcations/messages to the user
    /// </summary>
    public class MessageService : IMessageService
    {
        /// <summary>
        /// Shows a message with the given title and text.
        /// </summary>
        /// <param name="title">The title of the message</param>
        /// <param name="message">The text of the message</param>
        public void ShowInformation(string title, string message)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
