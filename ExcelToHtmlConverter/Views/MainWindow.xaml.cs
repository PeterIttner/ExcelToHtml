using ExcelToHtmlConverter.ViewModels;
using System.Windows;

namespace ExcelToHtmlConverter.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Creates a new instance of the <see cref="MainWindow"/> class.
        /// Also sets the DataContext of the View to a new instance of the <see cref="MainViewModel"/> class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            if (!System.ComponentModel.DesignerProperties.GetIsInDesignMode(this))
            {
                this.DataContext = new MainViewModel();
            }

        }
    }
}
