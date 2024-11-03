// ----------------------------------------------------------------------------------------
// Project Name  : ClarificationDetailsProject
// File Name     : MainWindow.xaml.cs
// Description   : Represents the mainwindow class
// Author        : Yahkoob P
// Date          : 27-10-2024
// ----------------------------------------------------------------------------------------

using System.Windows;
using ClarificationDetailsProject.ViewModels;

namespace ClarificationDetailsProject
{
    /// <summary>
    /// Main entry point for the application UI.
    /// Defines interaction logic for the main window and binds to the <see cref="ClarificationViewModel"/> for data operations.
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Gets or sets the ViewModel for managing data in the main window.
        /// </summary>
        private ClarificationViewModel ViewModel { get; set; } = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// Sets the DataContext to the <see cref="ClarificationViewModel"/>.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            ViewModel = new ClarificationViewModel();
            this.DataContext = ViewModel;
        }

        /// <summary>
        /// Handles the CheckBox Checked event and updates the selected modules in the ViewModel.
        /// </summary>
        /// <param name="sender">The CheckBox that was checked.</param>
        /// <param name="e">Event data for the Checked event.</param>
        private void ModulesChecked(object sender, RoutedEventArgs e)
        {
            ViewModel.UpdateSelectedModules();
        }

        /// <summary>
        /// Handles the CheckBox Unchecked event and updates the selected modules in the ViewModel.
        /// </summary>
        /// <param name="sender">The CheckBox that was unchecked.</param>
        /// <param name="e">Event data for the Unchecked event.</param>
        private void ModulesUnchecked(object sender, RoutedEventArgs e)
        {
            ViewModel.UpdateSelectedModules();
        }
    }
}
