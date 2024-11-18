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
        public void MergeFiles(ObservableCollection<string> pptPaths, string outputPath)
        {
            MessageBox.Show("Not implemented");
            return;
        }
    }
}
