// ----------------------------------------------------------------------------------------
// Project Name: PPTMerger
// File Name: MainWindow.cs
// Description: Code behind for MainWindow.xaml
// Author: Yahkoob P
// Date: 11-12-2024
// ----------------------------------------------------------------------------------------
using System.Windows;

namespace PPTMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //set the data context for main window
            this.DataContext = new PPTViewModel();
        }
    }
}
