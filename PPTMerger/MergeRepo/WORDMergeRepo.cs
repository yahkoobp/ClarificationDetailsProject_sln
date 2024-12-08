﻿using PPTMerger.Delegates;
using PPTMerger.Repo;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PPTMerger.MergeRepo
{
    internal class WORDMergeRepo : IRepo
    {
        public event EventHandler<FileProcessingFailedEventArgs> FileProcessingFailed;
        public event Action<string> LogEvent;
        public event EventHandler<int> ProgressEvent;
        protected void OnLog(string message)
        {
            LogEvent?.Invoke($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        }
        public async Task MergeFilesAsync(ObservableCollection<string> filePaths, string outputPath)
        {
            MessageBox.Show("Not implemented");
            return;
        }
    }
}
