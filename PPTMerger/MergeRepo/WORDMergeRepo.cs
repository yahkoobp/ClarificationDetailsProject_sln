using PPTMerger.Delegates;
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

        public void MergeFiles(ObservableCollection<string> filePaths, string outputPath)
        {
            MessageBox.Show("Not implemented");
            return;
        }
    }
}
