// ClarificationViewModel.cs
// 
// This file contains the implementation of the ClarificationViewModel class,
// which serves as a ViewModel for the Clarification Details feature in the 
// application. It inherits from ViewModelBase and is responsible for 
// managing the state and behavior of the user interface related to 
// clarifications.
// 
// Author: Yahkoob P
// Date: 2024-10-23

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClarificationDetailsProject.Commands;
using ClarificationDetailsProject.Models;
using System.Windows.Input;
using ClarificationDetailsProject.Repo;
using ClarificationDetailsProject.ExcelRepo;
using System.Windows;
using System.IO;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using System.Web.UI;
using System.Windows.Controls;

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
        private ObservableCollection<Clarification> _clarifications;
        private ObservableCollection<Models.Module> _modules;
        private Clarification _selectedClarification;
        public bool IsFilterApplied { get; set; } = false;
        public bool IsLoading { get; set; } = false;

        public ObservableCollection<Clarification> Clarifications
        {
            get { return _clarifications; }
            set
            {
                _clarifications = value;
                OnPropertyChanged(nameof(Clarifications));
               
            }
        }

        private ObservableCollection<Summary> summaries;
        public ObservableCollection<Summary> Summaries
        {
            get
            {
                return summaries;
            }
            set
            {
                summaries = value;
                OnPropertyChanged(nameof(Summaries));
            }
        }
        public ObservableCollection<Clarification> TempClarifications { get; set; }

        private ObservableCollection<Clarification> filteredClarifications;
        public ObservableCollection<Clarification> FilteredClarifications
        {
            get { return filteredClarifications; ; }
            set
            {
                filteredClarifications = value;
                OnPropertyChanged(nameof(FilteredClarifications));
            }
        }

        private string filePath;
        public string FilePath
        {
            get
            {
                return filePath;
            }
            set
            {
                filePath = value;
                OnPropertyChanged(nameof(FilePath));
            }
        }

        private string fileName;
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
                OnPropertyChanged(nameof(FileName));
            }
        }

        public ObservableCollection<Models.Module> Modules
        {
            get { return _modules; }
            set
            {
                _modules = value;
                OnPropertyChanged(nameof(Modules));
            }
        }


        private List<string> selectedModules;
        public List<string> SelectedModules
        {
            get { return selectedModules; }
            set
            {
                selectedModules = value;
                OnPropertyChanged(nameof(SelectedModules));
            }
        }

        private bool _isAllChecked;

        public bool IsAllChecked
        {
            get => _isAllChecked;
            set
            {
                _isAllChecked = value;
                OnPropertyChanged(nameof(IsAllChecked));

                foreach (var item in Modules)
                {
                    item.IsChecked = value;
                }
            }
        }

        private string _filterStatus;
        public string FilterStatus
        {
            get
            {
                return _filterStatus;
            }
            set
            {
                _filterStatus = value;
                OnPropertyChanged(nameof(FilterStatus));
              
            }
        }

        private DateTime filterFromDate;
        public DateTime FilterFromDate
        {
            get
            {
                return filterFromDate;
            }
            set
            {
                filterFromDate = value;
                OnPropertyChanged(nameof(FilterFromDate));

            }
        }

        private DateTime filterToDate;
        public DateTime FilterToDate
        {
            get
            {
                return filterToDate;
            }
            set
            {
                filterToDate = value;
                OnPropertyChanged(nameof(FilterToDate));

            }
        }

        private object _selectedTab;
        public object SelectedTab
        {
            get => _selectedTab;
            set
            {
                _selectedTab = value;
                OnPropertyChanged(nameof(SelectedTab));
            }
        }

        private string _searchText;
        public string SearchText
        {
            get 
            {
                return _searchText; 
            }
            set 
            {
                _searchText = value; 
                OnPropertyChanged(nameof(SearchText)); ApplyFilters(); 
            }
        }

        private string buttonText;
        public string ButtonText
        {
            get
            {
                return buttonText;
            }
            set
            {
                buttonText = value;
                OnPropertyChanged(nameof(ButtonText));
            }
        }

        // Command for loading Excel
        public ICommand LoadExcelCommand { get; }
        public ICommand ShowDialogCommand { get; }

        public ICommand ApplyFilterCommand { get; }

        public ICommand ResetFilterCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand SearchCommand {  get; }

        public ClarificationViewModel()
        {
            LoadExcelCommand = new RelayCommand(LoadExcelAsync);
            ShowDialogCommand = new RelayCommand(ShowDialog);
            ApplyFilterCommand = new RelayCommand(ApplyFilters);
            ResetFilterCommand = new RelayCommand(ResetFilter);
            ExportToExcelCommand = new RelayCommand(ExportToExcel);
            SearchCommand = new RelayCommand(Search);

            this.FilePath = string.Empty;
            Clarifications = new ObservableCollection<Clarification>()
            {
                 new Clarification { Number = 1, Date = DateTime.Now, DocumentName = "Doc1", Question = "Question1", Answer = "Answer1", Status = "Pending" },
            new Clarification{ Number = 2, Date = DateTime.Now, DocumentName = "Doc2", Question = "Question2", Answer = "Answer2", Status = "Closed" }
            };

            Summaries = new ObservableCollection<Summary>();
            filteredClarifications = new ObservableCollection<Clarification>();
            TempClarifications = new ObservableCollection<Clarification>();
            Modules = new ObservableCollection<Models.Module>();
            selectedModules = new List<string>();
            buttonText = "Show Details";
        }

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
        public async void LoadExcelAsync()
        {
            if (string.IsNullOrWhiteSpace(FilePath))
            {
                MessageBox.Show("Please select a file");
                return;
            }

            IsLoading = true;
            ButtonText = "Loading...";
            string filePath = this.FilePath;

            try
            {
                // Load data asynchronously
                var data = await _repo.LoadDataAsync(filePath);

                // Clear existing data and add the new data
                Clarifications.Clear();
                TempClarifications.Clear();
                foreach (var item in data)
                {
                    Clarifications.Add(item);
                    TempClarifications.Add(item);
                }

                // Load summaries and modules
                var summaries = _repo.GetSummaries();
                Summaries.Clear();
                foreach (var summary in summaries)
                {
                    Summaries.Add(summary);
                }

                var modules = _repo.GetModules();
                Modules.Clear();
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
        public void UpdateSelectedModules()
        {
            var SelectedModules = Modules.Where(m => m.IsChecked).ToList(); // Get selected modules
            this.SelectedModules.Clear();
            foreach (var module in SelectedModules)
            {
                this.SelectedModules.Add(module.Name);
            }
        }

        public void ApplyFilters()
        {
            IsFilterApplied = true;
            // Logic to filter based on _filterStatus and _searchText
            var filteredList = new ObservableCollection<Clarification>();
            FilteredClarifications.Clear();

            foreach (var clarification in TempClarifications)
            {
                bool matchesStatus = string.IsNullOrEmpty(FilterStatus) || clarification.Status.Equals(FilterStatus, StringComparison.OrdinalIgnoreCase) || FilterStatus.CompareTo("All") == 0;
                bool matchesModule = !selectedModules.Any() ||
                                     selectedModules.Contains(clarification.Module);
                //bool matchesDate = (FilterFromDate == null || clarification.Date >= FilterFromDate) &&
                //                   (FilterToDate == null || clarification.Date <= FilterToDate);

                // If all conditions are met, add the clarification to the filtered list
                if (matchesModule && matchesStatus)
                {
                    filteredList.Add(clarification);
                }

                foreach(var item in filteredList)
                {
                    FilteredClarifications.Add(item);
                }
            }
            Clarifications.Clear();
            foreach(var clarification in filteredList)
            {
                Clarifications.Add(clarification);
            }
        }

        public void Search()
        {
            if (string.IsNullOrWhiteSpace(SearchText))
                return; // Return if search text is empty

            SearchText = SearchText.ToLower();

            var results =  TempClarifications.Where(c =>
                c.Number.ToString().Contains(SearchText) ||
                (c.DocumentName?.ToLower().Contains(SearchText) ?? false) ||
                (c.Module?.ToLower().Contains(SearchText) ?? false) ||
                (c.Status?.ToLower().Contains(SearchText) ?? false) ||
                c.Date.ToString("yyyy-MM-dd").Contains(SearchText) ||
                (c.Question?.ToLower().Contains(SearchText) ?? false) ||
                (c.Answer?.ToLower().Contains(SearchText) ?? false)
            );

            Clarifications.Clear();
            foreach(var result in results)
            {
                Clarifications.Add(result);
            }
        }
        public void ResetFilter()
        {
            IsFilterApplied = false;
            filteredClarifications.Clear();
            Clarifications.Clear();
            foreach (var item in TempClarifications)
            {
                Clarifications.Add(item);
            }
        }

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
                    _repo.ExportClarificationsToExcel(FilteredClarifications, saveFileDialog.FileName);
                }
                else
                {
                    _repo.ExportClarificationsToExcel(Clarifications, saveFileDialog.FileName);
                }
                
            }
        }

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
                _repo.ExportSummaryToExcel(Summaries , saveFileDialog.FileName);

            }
        }


    }
}