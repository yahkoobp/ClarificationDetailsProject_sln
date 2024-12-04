using PPTMerger.Delegates;
using PPTMerger.Repo;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPTMerger.MergeRepo
{
    internal class ExcelMergeRepo : IRepo
    {
        public event EventHandler<FileProcessingFailedEventArgs> FileProcessingFailed;
        public event Action<string> LogEvent;
        protected void OnLog(string message)
        {
            LogEvent?.Invoke($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        }

        public void MergeFiles(ObservableCollection<string> pptPaths, string outputPath)
        {
            MessageBox.Show("Not implemented");
            return;
        }
    }
}
