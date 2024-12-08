﻿// ----------------------------------------------------------------------------------------
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
        //For storing paths for presentations
        private FileType selectedFileType;
        private string fileFilter;
        private string folderFilefilter;
        private bool isFolderSelection;
        private bool isFileSelection;
        private string selectedPath;
        private string mergeStatus;
        private ObservableCollection<string> selectedFiles;
        private ObservableCollection<string> logEntries;
        private IRepo repo;
        private int progressValue;
        private bool isMerging;
        private string progressStatus;

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
            repo.FileProcessingFailed += (sender, e) =>
            {
                MessageBox.Show($"Error processing file '{e.FilePath}': {e.ErrorMessage}");
            };
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

        public Dispatcher Dispatcher { get; }

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

        //Observable collection for storing selected file names
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

        public bool IsFolderSelection
        {
            get => isFolderSelection;
            set
            {
                isFolderSelection = value;
                OnPropertyChanged(nameof(IsFolderSelection));
            }
        }

        public bool IsFileSelection
        {
            get => !isFolderSelection;
            set
            {
                isFolderSelection = !value;
                OnPropertyChanged(nameof(IsFolderSelection));
            }
        }

        public string SelectedPath
        {
            get => selectedPath;
            set
            {
                selectedPath = value;
                OnPropertyChanged(nameof(IsFolderSelection));
            }
        }

        //Command for selecting files
        public ICommand SelectFilesCommand { get; }
        public ICommand RemoveItemCommand { get; }

        public ICommand ClearAllCommand { get; }
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

        public int ProgressValue
        {
            get => progressValue;
            set
            {
                progressValue = value;
                OnPropertyChanged(nameof(ProgressValue));
            }
        }      
        public bool IsMerging
        {
            get => isMerging;
            set
            {
                isMerging = value;
                OnPropertyChanged(nameof(IsMerging));
            }
        }

        public string ProgressStatus
        {
            get => progressStatus;
            set
            {
                progressStatus = value;
                OnPropertyChanged(nameof(ProgressStatus));
            }
        }

        //Command for merging presentations
        public ICommand MergeCommand { get; }

        private void UpdateFileTypeLogic(FileType fileType)
        {
            // Update logic or state based on the selected file type
            switch (fileType)
            {
                case FileType.PowerPoint:
                    repo = new PPTMergeRepo();
                    fileFilter = "PowerPoint Files (*.ppt;*.pptx)|*.ppt;*.pptx";
                    folderFilefilter = "*.pptx";
                    MergeStatus = "PowerPoint files...";
                    break;
                case FileType.PDF:
                    repo = new PDFMergeRepo();
                    fileFilter = "PDF Files (*.pdf)|*.pdf";
                    folderFilefilter = "*.pdf";
                    MergeStatus = "PDF files...";
                    break;
                case FileType.Excel:
                    repo = new ExcelMergeRepo();
                    fileFilter = "Excel Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv";
                    folderFilefilter = "*.xlsx";
                    MergeStatus = "Excel files...";
                    break;
                case FileType.Word:
                    repo = new WORDMergeRepo();
                    fileFilter = "Word Files (*.docx;*.doc)|*.docx;*.doc";
                    folderFilefilter = "*.docx";
                    MergeStatus = "Word files...";
                    break;
            }

            OnPropertyChanged(nameof(MergeStatus));
        }

        /// <summary>
        /// Function to select multiple presentations
        /// </summary>
        /// 
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
            }
        }

        private void LoadPresentationsFromFolder(string folderPath)
        {
            selectedFiles.Clear();
            var Files = Directory.GetFiles(folderPath, folderFilefilter, SearchOption.TopDirectoryOnly);
            foreach (var file in Files)
            {
                selectedFiles.Add(file);
            }

        }

        private void LoadPresentationsFromFiles(string[] files)
        {
            selectedFiles.Clear();
            foreach (var file in files)
            {
                selectedFiles.Add(file);
            }
        }

        private void RemoveFile(object file)
        {
            if (selectedFiles.Contains((string)file))
            {
                selectedFiles.Remove((string)file);
            }
        }

        private void ClearAllFiles(object obj)
        {
            SelectedFiles.Clear();
        }



        /// <summary>
        /// Function to call MergePowerPointFiles() function
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
                    ProgressValue = 0;
                    ProgressStatus = "Starting merge...";

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
                    IsMerging = false;
                }
            }
        }
    }
}
