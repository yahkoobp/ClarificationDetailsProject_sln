using PPTMerger.Repo;
using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace PPTMerger.MergeRepo
{
    internal class ExcelMergeRepo : IRepo
    {
        public event Action<string> LogEvent;
        public event EventHandler<int> ProgressEvent;
        protected void OnLog(string message)
        {
            LogEvent?.Invoke($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        }

        public async Task MergeFilesAsync(ObservableCollection<string> pptPaths, string outputPath)
        {
            //MessageBox.Show("Not implemented");
            return;
        }
    }
}
