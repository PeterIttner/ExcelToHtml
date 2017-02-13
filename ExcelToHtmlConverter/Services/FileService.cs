using System.IO;

namespace ExcelToHtmlConverter.Services
{
    /// <summary>
    /// Interface for file access abstraction.
    /// </summary>
    public interface IFileService
    {
        /// <summary>
        /// Writes the given content to the given file
        /// </summary>
        /// <param name="filename">The path to the file that should be written to.</param>
        /// <param name="content">The content that should be written</param>
        void WriteToFile(string filename, string content);
    }

    /// <summary>
    /// Handles access to files.
    /// </summary>
    public class FileService : IFileService
    {
        /// <summary>
        /// Writes the given content to the given file
        /// </summary>
        /// <param name="filename">The path to the file that should be written to.</param>
        /// <param name="content">The content that should be written</param>
        public void WriteToFile(string filename, string content)
        {
            File.WriteAllText(filename, content);
        }
    }
}
