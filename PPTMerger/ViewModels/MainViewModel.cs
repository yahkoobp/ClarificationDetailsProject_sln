// ----------------------------------------------------------------------------------------
// Project Name: PPTMerger
// File Name: PPTViewModel.cs
// Description: This file contains the implementation of the PPTViewModel class,
// which serves as a ViewModel for merging presentations in the 
// application. It inherits from ViewModelBase and is responsible for 
// managing the state and behavior of the user interface.
// Author: Yahkoob P
// Date: 11-12-2024
// ----------------------------------------------------------------------------------------
                                                                                
using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ClarificationDetailsProject.ViewModels;
using ClarificationDetailsProject.Commands;
using PPTMerger.Repo;
using PPTMerger.MergeRepo;
using System.IO;
using System.Windows.Forms;
using PPTMerger.Enums;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Linq;

namespace PPTMerger
{
    /// <summary>
    /// Represents the ViewModel for merging presentations
    /// </summary>
    /// <remarks>
    /// This class inherits from ViewModelBase and provides properties and 
    /// methods to manage the state of the user interface.
    /// It may include commands for user actions and properties for data binding.
    /// </remarks>
    public class MainViewModel : ViewModelBase
    {
        #region Constants
        private const string PPTFilter = "PowerPoint Files (*.ppt;*.pptx)|*.ppt;*.pptx";
        private const string PPTMergeStatus = "PowerPoint files...";
        private const string PDFFilter = "PDF Files (*.pdf)|*.pdf";
        private const string PDFMergeStatus = "PDF files...";
        private const string ExcelFilter = "Excel Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv";
        private const string ExcelMergeStatus = "Excel files...";
        private const string WordFilter = "Word Files (*.docx;*.doc)|*.docx;*.doc";
        private const string WordMergeStatus = "Word files...";
        #endregion

        #region Members
        private FileType selectedFileType = FileType.PowerPoint;
        private string fileFilter = string.Empty;
        private List<string> folderFilefilters = null;
        private bool isFolderSelection = false;
        private string selectedPath = string.Empty;
        private string mergeStatus = string.Empty;
        private ObservableCollection<string> selectedFiles = null;
        private ObservableCollection<string> logEntries = null;
        private IRepo repo = null;
        private int progressValue = 0;
        private bool isMerging = false;
        private string progressStatus = string.Empty;
        private bool isMergeButtonEnable = true;
        #endregion

        #region Constructor
        public MainViewModel()
        {
            selectedFiles = new ObservableCollection<string>();
            LogEntries = new ObservableCollection<string>();
            SelectedFileType = FileType.PowerPoint;
            SelectFilesCommand = new RelayCommand(SelectFiles);
            RemoveItemCommand = new RelayCommand(RemoveFile);
            MergeCommand = new RelayCommand(MergeFiles);
            ClearAllCommand = new RelayCommand(ClearAllFiles);
            repo = new PPTMergeRepo();
            Dispatcher = Dispatcher.CurrentDispatcher;

            repo.LogEvent += (string message) =>
            {
                Dispatcher.Invoke(() => LogEntries.Add(message));
            };
            repo.ProgressEvent += (sender, progress) =>
            {
                Dispatcher.Invoke(() => ProgressValue = progress);
            };
        }
        #endregion

        #region Properties
        /// <summary>
        /// For updating the UI from the background thread
        /// </summary>
        public Dispatcher Dispatcher { get; }

        /// <summary>
        /// Get or sets selected file type
        /// </summary>
        public FileType SelectedFileType
        {
            get => selectedFileType;
            set
            {
                selectedFileType = value;
                OnPropertyChanged(nameof(SelectedFileType));
                UpdateFileTypeLogic(value);
            }
        }

        /// <summary>
        /// Observable collection for storing selected file names
        /// </summary>
        public ObservableCollection<string> SelectedFiles
        {
            get
            {
                return selectedFiles;
            }
            set
            {
                selectedFiles = value;
                OnPropertyChanged(nameof(SelectedFiles));
            }
        }

        /// <summary>
        /// Collection to store log entries
        /// </summary>
        public ObservableCollection<string> LogEntries
        {
            get
            {
                return logEntries;
            }
            set
            {
                logEntries = value;
                OnPropertyChanged(nameof(LogEntries));
            }
        }
        /// <summary>
        /// To check wheather selected mode is folder
        /// </summary>
        public bool IsFolderSelection
        {
            get => isFolderSelection;
            set
            {
                isFolderSelection = value;
                OnPropertyChanged(nameof(IsFolderSelection));
            }
        }

        /// <summary>
        /// To check wheather selected mode is file
        /// </summary>
        public bool IsFileSelection
        {
            get => !isFolderSelection;
            set
            {
                isFolderSelection = !value;
                OnPropertyChanged(nameof(IsFolderSelection));
            }
        }
        /// <summary>
        /// Gets or sets the selected path
        /// </summary>
        public string SelectedPath
        {
            get => selectedPath;
            set
            {
                selectedPath = value;
                OnPropertyChanged(nameof(IsFolderSelection));
            }
        }

        public bool IsMergeButtonEnable
        {
            get
            {
                return isMergeButtonEnable;
            }
            set
            {
                isMergeButtonEnable = value;
                OnPropertyChanged(nameof(IsMergeButtonEnable));
            }
        }

        //Command for merging presentations
        public ICommand MergeCommand { get; }
        //Command for selecting files
        public ICommand SelectFilesCommand { get; }
        //Command for Remove an item from the selected files
        public ICommand RemoveItemCommand { get; }
        //Command for clearing selected files
        public ICommand ClearAllCommand { get; }

        /// <summary>
        /// Sets or gets the merge status
        /// </summary>
        public string MergeStatus {
            get
            {
                return mergeStatus;
            }
            set
            {
                mergeStatus = value;
                OnPropertyChanged(nameof(MergeStatus));
            }
        }

        /// <summary>
        /// Gets or sets the progress value
        /// </summary>
        public int ProgressValue
        {
            get => progressValue;
            set
            {
                progressValue = value;
                OnPropertyChanged(nameof(ProgressValue));
            }
        }

        /// <summary>
        /// To check wheather merging is on going or not
        /// </summary>
        public bool IsMerging
        {
            get => isMerging;
            set
            {
                isMerging = value;
                OnPropertyChanged(nameof(IsMerging));
            }
        }
        /// <summary>
        /// Gets or sets the progress status
        /// </summary>
        public string ProgressStatus
        {
            get => progressStatus;
            set
            {
                progressStatus = value;
                OnPropertyChanged(nameof(ProgressStatus));
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Method for updating logic for selected file type
        /// </summary>
        /// <param name="fileType"></param>
        private void UpdateFileTypeLogic(FileType fileType)
        {
            // Update logic or state based on the selected file type
            switch (fileType)
            {
                case FileType.PowerPoint:
                    repo = new PPTMergeRepo();
                    fileFilter = PPTFilter;
                    folderFilefilters = new List<string> { "*.ppt", "*pptx", "*ppsx" };
                    MergeStatus = PPTMergeStatus;
                    break; 
                case FileType.PDF:
                    repo = new PDFMergeRepo();
                    fileFilter = PDFFilter;
                    folderFilefilters = new List<string> { "*.pdf"};
                    MergeStatus = PDFMergeStatus;
                    break;
                case FileType.Excel:
                    repo = new ExcelMergeRepo();
                    fileFilter = ExcelFilter;
                    folderFilefilters = new List<string> { "*.xlsx"};
                    MergeStatus = ExcelMergeStatus;
                    break;
                case FileType.Word:
                    repo = new WORDMergeRepo();
                    fileFilter = WordFilter;
                    folderFilefilters = new List<string> { "*.docx" };
                    MergeStatus = WordMergeStatus;
                    break;
                default:
                    repo = null;
                    fileFilter = null;
                    folderFilefilters = null;
                    MergeStatus = null;
                    break;
            }

            OnPropertyChanged(nameof(MergeStatus));
        }

        /// <summary>
        /// Method to select multiple presentations
        /// </summary>
        private void SelectFiles(object obj)
        {
            if (IsFolderSelection)
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        SelectedPath = folderDialog.SelectedPath;
                        LoadPresentationsFromFolder(folderDialog.SelectedPath);
                    }
                    else
                    {
                        //ResetMembers();
                    }
                }
            }
            else
            {
                var fileDialog = new OpenFileDialog
                {
                    Filter = fileFilter,
                    Multiselect = true
                };

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    SelectedPath = string.Join("; ", fileDialog.FileNames);
                    LoadPresentationsFromFiles(fileDialog.FileNames);
                }
                else
                {
                    //ResetMembers();
                }
            }
        }

        /// <summary>
        /// Method to load presentations from folder
        /// </summary>
        /// <param name="folderPath"></param>
        private void LoadPresentationsFromFolder(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return;
            }
            selectedFiles.Clear();
            var Files = folderFilefilters.SelectMany(ext => Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly));
            foreach (var file in Files)
            {
                selectedFiles.Add(file);
            }

        }

        /// <summary>
        /// Method to load presentations from file
        /// </summary>
        /// <param name="files"></param>
        private void LoadPresentationsFromFiles(string[] files)
        {
            selectedFiles.Clear();
            foreach (var file in files)
            {
                selectedFiles.Add(file);
            }
        }

        /// <summary>
        /// Method to remove a file
        /// </summary>
        /// <param name="file"></param>
        private void RemoveFile(object file)
        {
            if (selectedFiles.Contains((string)file))
            {
                selectedFiles.Remove((string)file);
            }
        }

        /// <summary>
        /// Method to clear all files
        /// </summary>
        /// <param name="obj"></param>
        private void ClearAllFiles(object obj)
        {
            SelectedFiles.Clear();
        }

        /// <summary>
        /// Method to call MergePowerPointFiles() function
        /// </summary>
        private async void MergeFiles(object obj)
        {
            if (SelectedFiles.Count == 0)
            {
                System.Windows.MessageBox.Show($"No files selected.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = fileFilter,
                Title = "Save Merged files",
                FileName = "Merged"

            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    IsMerging = true;
                    IsMergeButtonEnable = false;
                    ProgressValue = 0;
                    ProgressStatus = "Merging...";
                    logEntries.Clear();

                    //Call MergePowerPointFiles() function 
                    await repo.MergeFilesAsync(selectedFiles, saveFileDialog.FileName);
                    ProgressStatus = "Merge completed successfully!";
                    System.Windows.MessageBox.Show($"Powerpoint presentations merged successfully.");
                }
                catch (Exception ex)
                {
                    ProgressStatus = $"Merge failed: {ex.Message}";
                    System.Windows.MessageBox.Show($"{ex.Message}");
                }
                finally
                {
                    IsMergeButtonEnable = true;
                    IsMerging = false;
                }
            }
            else
            {
                //ResetMembers();
            }
        }

        private void ResetMembers()
        {
            SelectedFileType = FileType.PowerPoint;
            IsFolderSelection = false;
            SelectedPath = string.Empty;
            MergeStatus = string.Empty;
            SelectedFiles = new ObservableCollection<string>();
            LogEntries = new ObservableCollection<string>();
            ProgressValue = 0;
            isMerging = false;
            progressStatus = string.Empty;
        }
        #endregion
    }
}
