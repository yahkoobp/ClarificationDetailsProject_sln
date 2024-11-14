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
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ClarificationDetailsProject.ViewModels;
using ClarificationDetailsProject.Commands;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Windows.Documents;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using System.Linq;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Threading;


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
    public class PPTViewModel : ViewModelBase
    {
        //For storing paths for presentations
        private ObservableCollection<string> selectedFiles;

        public PPTViewModel()
        {
            selectedFiles = new ObservableCollection<string>();
            SelectFilesCommand = new RelayCommand(SelectFiles);
            MergeCommand = new RelayCommand(MergePresentations);
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

        //Command for selecting files
        public ICommand SelectFilesCommand { get; }

        //Command for merging presentations
        public ICommand MergeCommand { get; }

        /// <summary>
        /// Function to select multiple presentations
        /// </summary>
        private void SelectFiles()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "PowerPoint Files|*.pptx;*.ppt"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                SelectedFiles.Clear();
                foreach (var file in openFileDialog.FileNames)
                {
                    SelectedFiles.Add(file);
                }
            }

        }

        /// <summary>
        /// Function to call MergePowerPointFiles() function
        /// </summary>
        private void MergePresentations()
        {
            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show($"No files selected.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PowerPoint Files|*.pptx",
                Title = "Save Merged Presentations",
                FileName = "MergedPresentations"

            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    //Call MergePowerPointFiles() function 
                    MergePresentations(selectedFiles, saveFileDialog.FileName);
                    MessageBox.Show($"Powerpoint presentations merged successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        /// <summary>
        /// Function to merge presentations
        /// </summary>
        /// <param name="paths"> Collection of path names </param>
        /// <param name="outPutPath">The output path that the merged presentation to be saved</param>
        private void MergePresentations(ObservableCollection<string> pptPaths, string outputPath)
        {
            var pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation mergedPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            try
            {
                foreach (string pptPath in pptPaths)
                {
                    Presentation sourcePresentation = pptApplication.Presentations.Open(
                        pptPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    try
                    {
                        foreach (Slide slide in sourcePresentation.Slides)
                        {
                            bool pasted = false;
                            int retryCount = 0;

                            while (!pasted && retryCount < 3)
                            {
                                try
                                {
                                    // Copy the slide to the clipboard.
                                    slide.Copy();
                                    Thread.Sleep(100);  // Short delay to allow copy operation to complete

                                    // Paste the slide into the merged presentation and retrieve it.
                                    Slide newSlide = mergedPresentation.Slides.Paste()[1];

                                    // Set the design and layout to match the source slide.
                                    newSlide.Design = slide.Design;
                                    newSlide.CustomLayout = slide.CustomLayout;

                                    pasted = true; // Successfully pasted
                                }
                                catch (COMException)
                                {
                                    retryCount++;
                                    Thread.Sleep(500); // Wait before retrying
                                }
                            }

                            if (!pasted)
                            {
                                throw new InvalidOperationException("Failed to paste slide after multiple attempts.");
                            }
                        }
                    }
                    finally
                    {
                        // Ensure the source presentation is closed after copying.
                        sourcePresentation.Close();
                        Marshal.ReleaseComObject(sourcePresentation);
                        Clipboard.Clear(); // Clear clipboard to reduce memory load
                    }
                }

                // Save the merged presentation to the output path.
                mergedPresentation.SaveAs(outputPath);
            }
            finally
            {
                // Ensure the merged presentation and PowerPoint application are closed and released.
                mergedPresentation.Close();
                pptApplication.Quit();
                Marshal.ReleaseComObject(mergedPresentation);
                Marshal.ReleaseComObject(pptApplication);
            }


        }
    }

}
