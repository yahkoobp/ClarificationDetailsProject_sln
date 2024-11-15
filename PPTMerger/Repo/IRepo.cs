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
        void MergePresentations(ObservableCollection<string> pptPaths, string outputPath);

    }
}
