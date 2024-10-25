using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClarificationDetailsProject.Models;

namespace ClarificationDetailsProject.Repo
{
    public interface IRepo
    {
        /// <summary>
        /// Loads the data from the data source
        /// </summary>
        /// <param name="filePath"></param>
        ObservableCollection<Clarification> LoadData(string filePath);
        /// <summary>
        /// Apply filters to the loaded data
        /// </summary>
        void ApplyFilters();
        /// <summary>
        /// Searches the data source for the specified text
        /// </summary>
        /// <param name="text"></param>
        void Search(string text);

        ObservableCollection<Models.Module> GetModules();

        ObservableCollection<Clarification> Filter(string status, DateTime? startDate, DateTime? endDate, List<string> selectedModuleNames);

        ObservableCollection<Summary> GetSummaries();

        void ExportClarificationsToExcel(string filename);
    }
}
