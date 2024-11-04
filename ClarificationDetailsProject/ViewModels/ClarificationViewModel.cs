// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: ClarificationViewModel.cs
// Description: This file contains the implementation of the ClarificationViewModel class,
// which serves as a ViewModel for the Clarification Details feature in the 
// application. It inherits from ViewModelBase and is responsible for 
// managing the state and behavior of the user interface related to 
// clarifications.
// Author: Yahkoob P
// Date: 27-10-2024
// ----------------------------------------------------------------------------------------
using ClarificationDetailsProject.Commands;
using ClarificationDetailsProject.ExcelRepo;
using ClarificationDetailsProject.Models;
using ClarificationDetailsProject.Repo;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ClarificationDetailsProject.ViewModels
{
    /// <summary>
    /// Represents the ViewModel for the Clarification Details feature.
    /// </summary>
    /// <remarks>
    /// This class inherits from ViewModelBase and provides properties and 
    /// methods to manage the state of the user interface for clarifications.
    /// It may include commands for user actions and properties for data binding.
    /// </remarks>
    public class ClarificationViewModel : ViewModelBase
    {
        //Constant to store successfull message
        private const string successfullMessageBoxText = "Exported successfully.";
        //Constant to store intial button text
        private const string initialButtonText = "Show Details";
        //Constant to store details tab text
        private const string detailsTab = "Details";
        //Constant to store summary tab text
        private const string summaryTab = "Summary";
        //Constant to store loading text
        private const string loadingText = "Loading...";

        private IRepo _repo = new ExcelDataRepo();

        /// <summary>
        /// Initializes a new instance of the <see cref="ClarificationViewModel"/> class.
        /// </summary>
        public ClarificationViewModel()
        {
            LoadExcelCommand = new RelayCommand(LoadExcelAsync);
            ShowDialogCommand = new RelayCommand(ShowDialog);
            ApplyFilterCommand = new RelayCommand(ApplyFilters);
            ResetFilterCommand = new RelayCommand(ResetFilter);
            ExportToExcelCommand = new RelayCommand(ExportToExcel);

            // Initialize properties with default values
            FilePath = string.Empty;
            Clarifications = new ObservableCollection<Clarification>();
            Summaries = new ObservableCollection<Summary>();
            FilteredClarifications = new ObservableCollection<Clarification>();
            FilterFromDate = null;
            FilterToDate = null;
            TempClarifications = new ObservableCollection<Clarification>();
            Modules = new ObservableCollection<Models.Module>();
            selectedModules = new List<string>();
            SelectedTab = new TabItem();
            SelectedTab.Header = detailsTab;
            //SelectedTab = new TabItem();
            buttonText = initialButtonText;
        }

        /// <summary>
        /// Indicates if a filter is currently applied.
        /// </summary>
        public bool IsFilterApplied { get; set; } = false;

        /// <summary>
        /// Indicates if a search is currently applied.
        /// </summary>
        public bool IsSearchApplied { get; set; } = false;

        /// <summary>
        /// Indicates if data is currently loading.
        /// </summary>
        public bool IsLoading { get; set; } = false;

        /// <summary>
        /// Collection of clarifications loaded from the data source.
        /// </summary>
        public ObservableCollection<Clarification> Clarifications
        {
            get { return clarifications; }
            set
            {
                clarifications = value;
                OnPropertyChanged(nameof(Clarifications));
            }
        }

        /// <summary>
        /// Temporary collection to hold clarifications for filtering purposes.
        /// </summary>
        public ObservableCollection<Clarification> TempClarifications { get; set; }

        /// <summary>
        /// Collection of summaries generated based on loaded data.
        /// </summary>
        public ObservableCollection<Summary> Summaries
        {
            get { return summaries; }
            set
            {
                summaries = value;
                OnPropertyChanged(nameof(Summaries));
            }
        }

        /// <summary>
        /// Collection of clarifications filtered based on user-selected criteria.
        /// </summary>
        public ObservableCollection<Clarification> FilteredClarifications
        {
            get { return filteredClarifications; }
            set
            {
                filteredClarifications = value;
                OnPropertyChanged(nameof(FilteredClarifications));
            }
        }

        /// <summary>
        /// Path to the selected file.
        /// </summary>
        public string FilePath
        {
            get { return filePath; }
            set
            {
                filePath = value;
                OnPropertyChanged(nameof(FilePath));
            }
        }

        /// <summary>
        /// Name of the selected file.
        /// </summary>
        public string FileName
        {
            get { return fileName; }
            set
            {
                fileName = value;
                OnPropertyChanged(nameof(FileName));
            }
        }

        /// <summary>
        /// Collection of modules available for filtering.
        /// </summary>
        public ObservableCollection<Models.Module> Modules
        {
            get { return modules; }
            set
            {
                modules = value;
                OnPropertyChanged(nameof(Modules));
            }
        }

        /// <summary>
        /// List of selected module names used for filtering.
        /// </summary>
        public List<string> SelectedModules
        {
            get { return selectedModules; }
            set
            {
                selectedModules = value;
                OnPropertyChanged(nameof(SelectedModules));
            }
        }

        /// <summary>
        /// Indicates whether all modules are selected.
        /// </summary>
        public bool IsAllChecked
        {
            get => isAllChecked;
            set
            {
                isAllChecked = value;
                OnPropertyChanged(nameof(IsAllChecked));

                foreach (var item in Modules)
                {
                    item.IsChecked = value;
                }
            }
        }

        /// <summary>
        /// Status filter applied to the clarifications.
        /// </summary>
        public string FilterStatus
        {
            get { return filterStatus; }
            set
            {
                filterStatus = value;
                OnPropertyChanged(nameof(FilterStatus));
            }
        }

        /// <summary>
        /// Start date for date filtering.
        /// </summary>
        public DateTime? FilterFromDate
        {
            get { return filterFromDate; }
            set
            {
                filterFromDate = value;
                OnPropertyChanged(nameof(FilterFromDate));
            }
        }

        /// <summary>
        /// End date for date filtering.
        /// </summary>
        public DateTime? FilterToDate
        {
            get { return filterToDate; }
            set
            {
                filterToDate = value;
                OnPropertyChanged(nameof(FilterToDate));
            }
        }

        /// <summary>
        /// Currently selected tab (Details or Summary).
        /// </summary>
        public TabItem SelectedTab
        {
            get => selectedTab;
            set
            {
                selectedTab = value;
                OnPropertyChanged(nameof(SelectedTab));
            }
        }

        /// <summary>
        /// Text input for searching clarifications.
        /// </summary>
        public string SearchText
        {
            get { return searchText; }
            set
            {
                searchText = value;
                OnPropertyChanged(nameof(SearchText));
                ApplyFilters();
            }
        }


        /// <summary>
        /// Text displayed on the action button.
        /// </summary>
        public string ButtonText
        {
            get { return buttonText; }
            set
            {
                buttonText = value;
                OnPropertyChanged(nameof(ButtonText));
            }
        }

        // Command for loading data from excel
        public ICommand LoadExcelCommand { get; }
        //Command to show the fileOpen dialog
        public ICommand ShowDialogCommand { get; }
        //Command to apply filters
        public ICommand ApplyFilterCommand { get; }
        //command to reset filters
        public ICommand ResetFilterCommand { get; }
        //commands to export data to excel
        public ICommand ExportToExcelCommand { get; }

        private ObservableCollection<Clarification> clarifications;
        private ObservableCollection<Models.Module> modules;
        private ObservableCollection<Summary> summaries;
        private ObservableCollection<Clarification> filteredClarifications;
        private string filePath;
        private string fileName;
        private List<string> selectedModules;      
        private bool isAllChecked;      
        private string filterStatus;
        private DateTime? filterFromDate;     
        private DateTime? filterToDate;
        private TabItem selectedTab;   
        private string searchText;
        private string buttonText;
       
        /// <summary>
        /// Opens a file dialog for selecting an Excel file.
        /// </summary>
        public void ShowDialog()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                this.FilePath = dialog.FileName;
                this.FileName = Path.GetFileName(FilePath);
            }
        }

        /// <summary>
        /// Loads clarifications data from the specified Excel file asynchronously.
        /// </summary>
        private async void LoadExcelAsync()
        {
            if (string.IsNullOrWhiteSpace(FilePath))
            {
                MessageBox.Show(messageBoxText: "Please select a file.",
                caption: "Alert",
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Warning);
                return;
            }

            IsLoading = true;
            ButtonText = loadingText;
            string filePath = this.FilePath;

            try
            {
                // Clear all collections before loading new data
                Clarifications.Clear();
                TempClarifications.Clear();
                Summaries.Clear();
                Modules.Clear();

                // Load data asynchronously
                var data = await _repo.LoadDataAsync(filePath);

                // Add new data to Clarifications and TempClarifications
                foreach (var item in data)
                {
                    Clarifications.Add(item);
                    TempClarifications.Add(item);
                }

                // Load summaries and modules
                var summaries = _repo.GetSummaries();
                foreach (var summary in summaries)
                {
                    Summaries.Add(summary);
                }

                var modules = _repo.GetModules();
                foreach (var item in modules)
                {
                    Modules.Add(item);
                }
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show($"Operation error: {ex.Message}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure loading state is reset
                IsLoading = false;
                ButtonText = initialButtonText;
            }
        }

        /// <summary>
        /// Updates the selected modules based on user selection.
        /// </summary>
        public void UpdateSelectedModules()
        {
            var SelectedModules = Modules.Where(m => m.IsChecked).ToList(); // Get selected modules
            this.SelectedModules.Clear();
            foreach (var module in SelectedModules)
            {
                this.SelectedModules.Add(module.Name);
            }
        }

        /// <summary>
        /// Applies filters based on user-selected criteria.
        /// </summary>
        private void ApplyFilters()
        {
            IsFilterApplied = true;
            IsSearchApplied = !string.IsNullOrWhiteSpace(SearchText); // Check if search is needed

            var filteredList = new ObservableCollection<Clarification>();
            FilteredClarifications.Clear();

            // Convert SearchText to lowercase for case-insensitive search if applicable
            var searchTextLower = SearchText?.ToLower();

            foreach (var clarification in TempClarifications)
            {
                // Check status and module filters
                bool matchesStatus = string.IsNullOrEmpty(FilterStatus) ||
                                     clarification.Status.Equals(FilterStatus, StringComparison.OrdinalIgnoreCase) ||
                                     FilterStatus.Equals("All", StringComparison.OrdinalIgnoreCase);
                bool matchesModule = !selectedModules.Any() || selectedModules.Contains(clarification.Module);
                bool matchesDate = (FilterFromDate == null || clarification.Date >= FilterFromDate) &&
                                   (FilterToDate == null || clarification.Date <= FilterToDate);

                // Check if it matches search criteria if SearchText is not empty
                bool matchesSearch = true;
                if (IsSearchApplied)
                {
                    matchesSearch = clarification.Number.ToString().Contains(searchTextLower) ||
                                    (clarification.DocumentName?.ToLower().Contains(searchTextLower) ?? false) ||
                                    (clarification.Module?.ToLower().Contains(searchTextLower) ?? false) ||
                                    (clarification.Status?.ToLower().Contains(searchTextLower) ?? false) ||
                                    clarification.Date.ToString("yyyy-MM-dd").Contains(searchTextLower) ||
                                    (clarification.Question?.ToLower().Contains(searchTextLower) ?? false) ||
                                    (clarification.Answer?.ToLower().Contains(searchTextLower) ?? false);
                }

                // If all conditions are met, add the clarification to the filtered list
                if (matchesStatus && matchesModule && matchesDate && matchesSearch)
                {
                    filteredList.Add(clarification);
                }
            }

            // Update the collections
            foreach (var item in filteredList)
            {
                FilteredClarifications.Add(item);
            }

            Clarifications.Clear();
            foreach (var clarification in filteredList)
            {
                Clarifications.Add(clarification);
            }
        }

        /// <summary>
        /// Resets all filters and clears filtered data.
        /// </summary>
        private void ResetFilter()
        {
            IsFilterApplied = false;
            IsSearchApplied = false;
            IsAllChecked = false;            
            FilterFromDate = null;
            FilterToDate = null;
            FilterStatus = null;
            FilteredClarifications.Clear();
            Clarifications.Clear();
            foreach (var item in TempClarifications)
            {
                Clarifications.Add(item);
            }
            SearchText = null;
        }

        /// <summary>
        /// Exports data to Excel based on the selected tab.
        /// </summary>
        private void ExportToExcel()
        {
            if (SelectedTab is TabItem tabItem)
            {
                if (tabItem.Header.ToString() == detailsTab)
                {
                    ExportClarificationToExcel();
                }
                else if (tabItem.Header.ToString() == summaryTab)
                {
                    ExportSummaryToExcel();
                }
            }
        }

        /// <summary>
        /// Exports the clarifications data to an Excel file.
        /// </summary>
        private void ExportClarificationToExcel()
        {
            //check if there is no clarifications
            if(filteredClarifications.Count == 0 || clarifications.Count == 0)
            {
                MessageBox.Show(messageBoxText: "No clarifications to export.",
                caption: "Alert",
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Warning);
                return;
            }
            // Open a SaveFileDialog to specify the file path
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Save Clarifications File",
                FileName = "Clarifications.xlsx" // Default file name
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                if (IsFilterApplied)
                {
                    try
                    {
                        _repo.ExportClarificationsToExcel(FilteredClarifications, saveFileDialog.FileName);
                        MessageBox.Show(successfullMessageBoxText ,
                        caption: "Success",
                        button: MessageBoxButton.OK,
                        icon: MessageBoxImage.Information);
                    }
                    catch (COMException ex)
                    {

                        MessageBox.Show($"{ex.Message}");
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        MessageBox.Show($"{ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{ex.Message}");
                    }

                }
                else
                {
                    try
                    {
                        _repo.ExportClarificationsToExcel(Clarifications, saveFileDialog.FileName);
                        MessageBox.Show(successfullMessageBoxText,
                        caption: "Success",
                        button: MessageBoxButton.OK,
                        icon: MessageBoxImage.Information);
                    }
                    catch (COMException ex)
                    {

                        MessageBox.Show($"{ex.Message}");
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        MessageBox.Show($"{ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{ex.Message}");
                    }

                }
                
            }
        }

        /// <summary>
        /// Exports the summaries data to an Excel file.
        /// </summary>
        private void ExportSummaryToExcel()
        {
            //check if there is no summmaries to export
            if (Summaries.Count == 0)
            {
                MessageBox.Show(messageBoxText: "No summaries to export.",
                caption: "Alert",
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Warning);
                return;
            }
            // Open a SaveFileDialog to specify the file path
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Save Summary File",
                FileName = "Summaries.xlsx" // Default file name
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    _repo.ExportSummaryToExcel(Summaries, saveFileDialog.FileName);
                    MessageBox.Show(successfullMessageBoxText,
                    caption: "Success",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Information);
                }
                catch (COMException ex)
                {

                    MessageBox.Show($"{ex.Message}");
                }
                catch (UnauthorizedAccessException ex)
                {
                    MessageBox.Show($"{ex.Message}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"{ex.Message}");
                }
            }
        }

    }
}