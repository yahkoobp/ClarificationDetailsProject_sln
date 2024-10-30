using System.Windows;
using ClarificationDetailsProject.ViewModels;

namespace ClarificationDetailsProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ClarificationViewModel ViewModel { get; set; } = null;      
        public MainWindow()
        {
            InitializeComponent();
            ViewModel = new ClarificationViewModel();
            this.DataContext = ViewModel;
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            ViewModel.UpdateSelectedModules();
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            ViewModel.UpdateSelectedModules();
        }
    }
}
