// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: IRepo.cs
// Description: Defines an interface for managing data operations in the ClarificationDetailsProject.
// Author: Yahkoob P
// Date: 27-10-2024
// ----------------------------------------------------------------------------------------

using System.Collections.ObjectModel;
using System.Threading.Tasks;
using ClarificationDetailsProject.Models;

namespace ClarificationDetailsProject.Repo
{
    /// <summary>
    /// Provides a contract for handling data operations in the ClarificationDetailsProject.
    /// </summary>
    /// <remarks>
    /// The IRepo interface defines the necessary methods for loading, searching, retrieving modules,
    /// generating summaries, and exporting data to Excel. This interface facilitates dependency 
    /// injection and promotes modular, testable code.
    /// </remarks>
    public interface IRepo
    {
        /// <summary>
        /// Asynchronously loads clarification data from the specified file path.
        /// </summary>
        /// <param name="filePath">The file path of the data source.</param>
        /// <returns>A task that represents the asynchronous operation. The task result contains
        /// an ObservableCollection of <see cref="Clarification"/> objects.</returns>
        Task<ObservableCollection<Clarification>> LoadDataAsync(string filePath);

        /// <summary>
        /// Retrieves the list of modules from the data source.
        /// </summary>
        /// <returns>An ObservableCollection of <see cref="Models.Module"/> objects.</returns>
        ObservableCollection<Models.Module> GetModules();

        /// <summary>
        /// Generates summary data grouped by module.
        /// </summary>
        /// <returns>An ObservableCollection of <see cref="Summary"/> objects representing
        /// summary statistics by module.</returns>
        ObservableCollection<Summary> GetSummaries();

        /// <summary>
        /// Exports a collection of clarifications to an Excel file.
        /// </summary>
        /// <param name="clarifications">The collection of <see cref="Clarification"/> objects to export.</param>
        /// <param name="filename">The file name for saving the Excel document.</param>
        void ExportClarificationsToExcel(ObservableCollection<Clarification> clarifications, string filename);

        /// <summary>
        /// Exports a collection of summaries to an Excel file.
        /// </summary>
        /// <param name="summaries">The collection of <see cref="Summary"/> objects to export.</param>
        /// <param name="filename">The file name for saving the Excel document.</param>
        void ExportSummaryToExcel(ObservableCollection<Summary> summaries, string filename);
    }
}

