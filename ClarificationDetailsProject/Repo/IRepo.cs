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
        Task<ObservableCollection<Clarification>> LoadDataAsync(string filePath);
        /// <summary>
        /// Searches the data source for the specified text
        /// </summary>
        /// <param name="text"></param>
        void Search(string text);
        ObservableCollection<Models.Module> GetModules();
        ObservableCollection<Summary> GetSummaries();
        void ExportClarificationsToExcel(ObservableCollection<Clarification> clarifications , string filename);
        void ExportSummaryToExcel(ObservableCollection<Summary> summaries, string filename);


    }
}
