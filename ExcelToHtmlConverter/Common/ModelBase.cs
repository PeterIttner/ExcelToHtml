using System.ComponentModel;

namespace ExcelToHtmlConverter.Common
{
    /// <summary>
    /// Implements the INotifyPropertyChanged interface for the MVVM concept.
    /// </summary>
    public class ModelBase : INotifyPropertyChanged
    {
        /// <summary>
        /// Fired when a MVVM property has been changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Fires the PropertyChanged event with the given name as property.
        /// </summary>
        /// <param name="property">The name of the property that has been changed.</param>
        public void NotifyPropertyChanged(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}
