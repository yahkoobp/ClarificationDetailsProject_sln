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
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ClarificationDetailsProject.ViewModels;
using ClarificationDetailsProject.Commands;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Windows.Documents;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
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
                    MergePowerPointFiles(selectedFiles, saveFileDialog.FileName);
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
        public void MergePowerPointFiles(ObservableCollection<string> pptPaths, string outputPath)
        {
            using (PresentationDocument mergedPresentation = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation))
            {
                // Create a new PresentationPart for the merged presentation
                PresentationPart mergedPresentationPart = mergedPresentation.AddPresentationPart();
                mergedPresentationPart.Presentation = new Presentation();
                mergedPresentationPart.Presentation.SlideIdList = new SlideIdList();

                // A dictionary to track already copied SlideMasterParts and SlideLayoutParts
                Dictionary<string, SlideMasterPart> slideMasterParts = new Dictionary<string, SlideMasterPart>();
                Dictionary<string, SlideLayoutPart> slideLayoutParts = new Dictionary<string, SlideLayoutPart>();

                foreach (string pptPath in pptPaths)
                {
                    using (PresentationDocument sourcePresentation = PresentationDocument.Open(pptPath, false))
                    {
                        PresentationPart sourcePresentationPart = sourcePresentation.PresentationPart;

                        // Copy SlideMasters to merged presentation
                        foreach (SlideMasterPart masterPart in sourcePresentationPart.SlideMasterParts)
                        {
                            // Check if the SlideMasterPart is already added to avoid duplicates
                            if (!slideMasterParts.ContainsKey(masterPart.Uri.ToString()))
                            {
                                SlideMasterPart newMasterPart = mergedPresentationPart.AddNewPart<SlideMasterPart>();
                                newMasterPart.FeedData(masterPart.GetStream()); // Copy master content
                                slideMasterParts[masterPart.Uri.ToString()] = newMasterPart;

                                // Copy SlideLayouts associated with the SlideMasterPart
                                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts)
                                {
                                    // Check if layout is already copied
                                    if (!slideLayoutParts.ContainsKey(layoutPart.Uri.ToString()))
                                    {
                                        SlideLayoutPart newLayoutPart = newMasterPart.AddNewPart<SlideLayoutPart>();
                                        newLayoutPart.FeedData(layoutPart.GetStream());
                                        slideLayoutParts[layoutPart.Uri.ToString()] = newLayoutPart;
                                    }
                                }
                            }
                        }

                        // Add slides from source presentation
                        foreach (SlideId slideId in sourcePresentationPart.Presentation.SlideIdList)
                        {
                            SlidePart sourceSlidePart = (SlidePart)sourcePresentationPart.GetPartById(slideId.RelationshipId);

                            // Create a new SlidePart in the merged presentation and copy the slide content
                            SlidePart newSlidePart = mergedPresentationPart.AddNewPart<SlidePart>();
                            newSlidePart.FeedData(sourceSlidePart.GetStream());

                            // Add the slide to the merged presentation's SlideIdList with a new unique ID
                            SlideId newSlideId = new SlideId
                            {
                                Id = (UInt32Value)(mergedPresentationPart.Presentation.SlideIdList.ChildElements.Count + 256U),
                                RelationshipId = mergedPresentationPart.GetIdOfPart(newSlidePart)
                            };
                            mergedPresentationPart.Presentation.SlideIdList.Append(newSlideId);
                        }
                    }
                }

                // Save the merged presentation
                mergedPresentationPart.Presentation.Save();
            }
        }
    }
}