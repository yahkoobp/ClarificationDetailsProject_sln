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
        private string _filterStatus;
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

        public ObservableCollection<Models.Module> Modules
        {
            get { return _modules; }
            set
            {
                _modules = value;
                OnPropertyChanged(nameof(Modules));
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



        public string FilterStatus
        {
            get { return _filterStatus; }
            set { _filterStatus = value; OnPropertyChanged(nameof(FilterStatus)); ApplyFilters(); }
        }

        public string SearchText
        {
            get { return _searchText; }
            set { _searchText = value; OnPropertyChanged(nameof(SearchText)); ApplyFilters(); }
        }

        // Command for loading Excel
        public ICommand LoadExcelCommand { get; }

        public ClarificationViewModel()
        {
            LoadExcelCommand = new RelayCommand(LoadExcel);
            Clarifications = new ObservableCollection<Clarification>();
            Modules = new ObservableCollection<Models.Module>();
        }

        private void LoadExcel()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

            if (dialog.ShowDialog() == true)
            {
                string filePath = dialog.FileName;
                try
                {
                    var data = _repo.LoadData(filePath);
                    Clarifications.Clear();

                    foreach (var item in data)
                    {
                        Clarifications.Add(item);
                    }

                    var modules = _repo.GetModules();
                    Modules.Clear();
                    foreach(var item in modules)
                    {
                        Modules.Add(item);
                    }
                }
                catch(InvalidOperationException ex)
                {
                    MessageBox.Show($"{ex.Message}");
                }
                catch(Exception ex)
                {
                    MessageBox.Show($"{ex.Message}");
                }         
                
            }
        }

        private void ApplyFilters()
        {
            // Logic to filter based on _filterStatus and _searchText
            var filtered = Clarifications
                            .Where(c => (string.IsNullOrEmpty(FilterStatus) || c.Status == FilterStatus) &&
                                        (string.IsNullOrEmpty(SearchText) || c.Question.Contains(SearchText)))
                            .ToList();

            Clarifications = new ObservableCollection<Clarification>(filtered);
        }
    }
}