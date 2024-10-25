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
        private string _searchText;

        public ObservableCollection<Clarification> Clarifications
        {
            get { return _clarifications; }
            set
            {
                _clarifications = value;
                OnPropertyChanged(nameof(Clarifications));
               
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
        public string SearchText
        {
            get { return _searchText; }
            set { _searchText = value; OnPropertyChanged(nameof(SearchText)); ApplyFilters(); }
        }

        // Command for loading Excel
        public ICommand LoadExcelCommand { get; }
        public ICommand ShowDialogCommand { get; }

        public ICommand ApplyFilterCommand { get; }

        public ICommand ResetFilterCommand { get; }

        public ClarificationViewModel()
        {
            LoadExcelCommand = new RelayCommand(LoadExcel);
            ShowDialogCommand = new RelayCommand(ShowDialog);
            ApplyFilterCommand = new RelayCommand(ApplyFilters);
            ResetFilterCommand = new RelayCommand(ResetFilter);
            this.FilePath = string.Empty;
            Clarifications = new ObservableCollection<Clarification>()
            {
                 new Clarification { Number = 1, Date = DateTime.Now, DocumentName = "Doc1", Question = "Question1", Answer = "Answer1", Status = "Pending" },
            new Clarification{ Number = 2, Date = DateTime.Now, DocumentName = "Doc2", Question = "Question2", Answer = "Answer2", Status = "Closed" }
            };
            filteredClarifications = new ObservableCollection<Clarification>();
            TempClarifications = new ObservableCollection<Clarification>();
            Modules = new ObservableCollection<Models.Module>();
            selectedModules = new List<string>();
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
        public void LoadExcel()
        {
            string filePath = this.FilePath;
            try
            {
                var data = _repo.LoadData(filePath);
                Clarifications.Clear();
                TempClarifications.Clear();
                foreach (var item in data)
                {
                    Clarifications.Add(item);
                    TempClarifications.Add(item);
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
                MessageBox.Show($"{ex.Message}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
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
            // Logic to filter based on _filterStatus and _searchText
            var filteredList = new ObservableCollection<Clarification>();
            FilteredClarifications.Clear();

            foreach (var clarification in TempClarifications)
            {
                bool matchesStatus = string.IsNullOrEmpty(FilterStatus) || clarification.Status.Equals(FilterStatus, StringComparison.OrdinalIgnoreCase);
                bool matchesModule = !selectedModules.Any() ||
                                     selectedModules.Contains(clarification.Module);

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

        public void ResetFilter()
        {
            Clarifications.Clear();
            foreach (var item in TempClarifications)
            {
                Clarifications.Add(item);
            }
        }
    }
}