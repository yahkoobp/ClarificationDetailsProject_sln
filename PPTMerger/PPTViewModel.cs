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
using Microsoft.Office.Interop.PowerPoint;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ClarificationDetailsProject.ViewModels;
using ClarificationDetailsProject.Commands;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

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

            if(openFileDialog.ShowDialog() == true)
            {
                SelectedFiles.Clear();
                foreach(var file in openFileDialog.FileNames)
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
            if(SelectedFiles.Count == 0)
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
                    MergePowerPointFiles(selectedFiles, saveFileDialog.FileName);
                    MessageBox.Show($"Powerpoint presentations merged successfully.");
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message );
                }
            }
        }

        /// <summary>
        /// Function to merge presentations
        /// </summary>
        /// <param name="paths"> Collection of path names </param>
        /// <param name="outPutPath">The output path that the merged presentation to be saved</param>
        private void MergePowerPointFiles(ObservableCollection<string> paths , string outPutPath)
        {
            //Create a new presentation application instance
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            //Add a new presenation to serve as the merged presenation
            Presentation mergedPresentaion = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            try
            {
                foreach(string path in paths)
                {
                    //Open each presenation files in readonly mode
                    Presentation currentPresentation = pptApplication.Presentations.Open(path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    //For each slide in the presentation file do the following
                    foreach(Slide slide in currentPresentation.Slides)
                    {
                        //Copy each slide from the presentation
                        slide.Copy();
                        //Paste the slide in to the merged presentation at the end
                        mergedPresentaion.Slides.Paste(mergedPresentaion.Slides.Count + 1);
                    }
                    //Close the source presentation after copying its slides 
                    currentPresentation.Close();
                    Marshal.ReleaseComObject(currentPresentation);
                }
                //Save the merged presentations
                mergedPresentaion.SaveAs(outPutPath);
            }
            catch(Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
            finally
            {
                //Cleanup the resources
                mergedPresentaion.Close();
                pptApplication.Quit();
                Marshal.ReleaseComObject(mergedPresentaion);
                Marshal.ReleaseComObject(pptApplication);
            }

        }

    }
}
