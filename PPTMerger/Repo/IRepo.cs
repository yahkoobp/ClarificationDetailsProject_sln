using PPTMerger.Delegates;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTMerger.Repo
{
    internal interface IRepo
    {
        event EventHandler<FileProcessingFailedEventArgs> FileProcessingFailed;
        event Action<string> LogEvent;
        void MergeFiles(ObservableCollection<string> filePaths, string outputPath);

    }
}
