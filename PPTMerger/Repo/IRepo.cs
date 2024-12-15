// ----------------------------------------------------------------------------------------
// Project Name: PPTMerger
// File Name: IRepo.cs
// Description:Act as a repo for file merge class
// Author: Yahkoob P
// Date: 11-12-2024
// ----------------------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace PPTMerger.Repo
{
    internal interface IRepo
    {
        /// <summary>
        /// Event for logger
        /// </summary>
        event Action<string> LogEvent;
        /// <summary>
        /// Event for to handle merge progress
        /// </summary>
        event EventHandler<int> ProgressEvent;
        /// <summary>
        /// Method to merge files
        /// </summary>
        /// <param name="filePaths"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        Task MergeFilesAsync(ObservableCollection<string> filePaths, string outputPath);
    }
}
