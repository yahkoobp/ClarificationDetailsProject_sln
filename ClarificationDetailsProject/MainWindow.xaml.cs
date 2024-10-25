using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClarificationDetailsProject.ViewModels;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

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
