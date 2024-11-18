using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTMerger.Delegates
{
    public class FileProcessingFailedEventArgs : EventArgs
    {
        public string FilePath { get; }
        public string ErrorMessage { get; }

        public FileProcessingFailedEventArgs(string filePath, string errorMessage)
        {
            FilePath = filePath;
            ErrorMessage = errorMessage;
        }
    }
}
