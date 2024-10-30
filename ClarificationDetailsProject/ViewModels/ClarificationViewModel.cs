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
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ClarificationDetailsProject.Commands;
using ClarificationDetailsProject.Models;
using System.Windows.Input;
using ClarificationDetailsProject.Repo;
using ClarificationDetailsProject.ExcelRepo;
using System.Windows;
using System.IO;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Runtime.InteropServices;

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
        private IRepo _repo = new ExcelDataRepo();
        private ObservableCollection<Clarification> clarifications;
        private ObservableCollection<Models.Module> modules;
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

        private ObservableCollection<Summary> summaries;

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
        /// Temporary collection to hold clarifications for filtering purposes.
        /// </summary>
        public ObservableCollection<Clarification> TempClarifications { get; set; }

        private ObservableCollection<Clarification> filteredClarifications;

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

        private string filePath;

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

        private string fileName;

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

        private List<string> selectedModules;

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

        private bool isAllChecked;

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

        private string filterStatus;

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

        private DateTime? filterFromDate;

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

        private DateTime? filterToDate;

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

        private object _selectedTab;

        /// <summary>
        /// Currently selected tab (Details or Summary).
        /// </summary>
        public object SelectedTab
        {
            get => _selectedTab;
            set
            {
                _selectedTab = value;
                OnPropertyChanged(nameof(SelectedTab));
            }
        }

        private string searchText;

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

        private string buttonText;

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

        // Commands for various actions
        public ICommand LoadExcelCommand { get; }
        public ICommand ShowDialogCommand { get; }
        public ICommand ApplyFilterCommand { get; }
        public ICommand ResetFilterCommand { get; }
        public ICommand ExportToExcelCommand { get; }

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
            Clarifications = new ObservableCollection<Clarification>
            {
            new Clarification { Number = 1, Date = DateTime.Now, DocumentName = "Doc1", Question = "Question1", Answer = "Answer1", Status = "Pending" },
            new Clarification{ Number = 2, Date = DateTime.Now, DocumentName = "Doc2", Question = "Question2", Answer = "Answer2", Status = "Closed" }
            };
            Summaries = new ObservableCollection<Summary>();
            FilteredClarifications = new ObservableCollection<Clarification>();
            FilterFromDate = null;
            FilterToDate = null;
            TempClarifications = new ObservableCollection<Clarification>();
            Modules = new ObservableCollection<Models.Module>();
            selectedModules = new List<string>();
            SelectedTab = "Details";
            buttonText = "Show Details";
        }

        /// <summary>
        /// Opens a file dialog for selecting an Excel file.
        /// </summary>

        public void ShowDialog()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

            if (dialog.ShowDialog() == true)
            {
                this.FilePath = dialog.FileName;
                this.FileName = Path.GetFileName(FilePath);
            }
        }


        /// <summary>
        /// Loads clarifications data from the specified Excel file asynchronously.
        /// </summary>
        public async void LoadExcelAsync()
        {
            if (string.IsNullOrWhiteSpace(FilePath))
            {
                MessageBox.Show(messageBoxText: "Please select a file",
                caption: "Alert",
                button: MessageBoxButton.OK,
                icon: MessageBoxImage.Warning);
                return;
            }

            IsLoading = true;
            ButtonText = "Loading...";
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
                ButtonText = "Show Details";
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
        public void ApplyFilters()
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
        public void ResetFilter()
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
        public void ExportToExcel()
        {
            if (SelectedTab is TabItem tabItem)
            {
                if (tabItem.Header.ToString() == "Details")
                {
                    ExportClarificationToExcel();
                }
                else if (tabItem.Header.ToString() == "Summary")
                {
                    ExportSummaryToExcel();
                }
            }
        }

        /// <summary>
        /// Exports the clarifications data to an Excel file.
        /// </summary>
        public void ExportClarificationToExcel()
        {
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
                        MessageBox.Show($"Exported Successfully");
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
                        MessageBox.Show($"Exported Successfully");
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
        public void ExportSummaryToExcel()
        {
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
                    MessageBox.Show($"Exported Successfully");
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