﻿// ----------------------------------------------------------------------------------------
// Project Name  : ClarificationDetailsProject
// File Name     : ExcelDataRepo.cs
// Description   : Represents a repository for handling Excel data operations.
// Author        : Yahkoob P
// Date          : 27-10-2024
// ----------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using ClarificationDetailsProject.Models;
using ClarificationDetailsProject.Repo;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClarificationDetailsProject.ExcelRepo
{
    /// <summary>
    /// Represents a repository for handling Excel data operations.
    /// </summary>
    /// <remarks>
    /// This class implements the IRepo interface and provides methods to
    /// load, filter, and search data within Excel files. It is designed
    /// to encapsulate all Excel-related data access logic.
    /// </remarks>
    public class ExcelDataRepo : IRepo
    {
        private ObservableCollection<Clarification> Clarifications = new ObservableCollection<Clarification>();
        private ObservableCollection<Models.Module> Modules = new ObservableCollection<Models.Module>();
        private readonly List<string> expectedHeaders = new List<string>() 
        {
            "No", 
            "Date",
            "Document Name and its section",
            "Page No",
            "Section Number",
            "Question",
            "Due Date",
            "Answer",
            "Priority",
            "status",
            "Remarks"
        };

        /// <summary>
        /// Validates the headers of an Excel worksheet against the expected headers.
        /// </summary>
        /// <param name="worksheet">The worksheet to validate.</param>
        /// <returns>True if the headers match; otherwise, false.</returns>
        public bool IsValidExcelWorkBook(Excel.Worksheet worksheet)
        {
            try
            {
                string sheetName = worksheet.Name;
                List<string> actualHeaders = new List<string>();

                // Read the second row (headers) in the sheet
                Excel.Range headerRange = worksheet.Range["A2",worksheet.Cells[2, expectedHeaders.Count]];
                foreach (Excel.Range cell in headerRange)
                {
                    actualHeaders.Add(cell.Value?.ToString() ?? string.Empty);
                }

                // Compare actual headers with expected headers
                if (!HeadersMatch(expectedHeaders, actualHeaders))
                {
                    return false;
                    // Handle the invalid sheet (log, skip, notify, etc.)
                }
                else
                {
                    return true;
                    // Process the valid sheet
                }

            }
            catch (Exception ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Checks if the actual headers match the expected headers.
        /// </summary>
        /// <param name="expectedHeaders">The list of expected headers.</param>
        /// <param name="actualHeaders">The list of actual headers from the worksheet.</param>
        /// <returns>True if headers match; otherwise, false.</returns>
        private bool HeadersMatch(List<string> expectedHeaders, List<string> actualHeaders)
        {
            if (expectedHeaders.Count != actualHeaders.Count)
                return false;

            for (int i = 0; i < expectedHeaders.Count; i++)
            {
                if (!expectedHeaders[i].Equals(actualHeaders[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Asynchronously loads data from an Excel file into the Clarifications collection.
        /// </summary>
        /// <param name="filePath">The file path of the Excel document to load.</param>
        /// <returns>A collection of Clarification objects loaded from the Excel file.</returns>
        public async Task<ObservableCollection<Clarification>> LoadDataAsync(string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;
            Modules.Clear();

            try
            {
                //Clear previous data
                Clarifications.Clear();
                // Open the workbook asynchronously
                workbook = await Task.Run(() => excelApp.Workbooks.Open(filePath));

                // Check if the workbook has any worksheets
                if (workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("The Excel file does not contain any worksheets.");
                }

                // Loop through each worksheet in the workbook
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    Excel.Range usedRange = worksheet.UsedRange;
                    int rowCount = usedRange.Rows.Count;

                    // Validate the worksheet headers
                    if (IsValidExcelWorkBook(worksheet))
                    {
                        // Add valid sheet names to modules list
                        Modules.Add(new Models.Module()
                        {
                            Name = worksheet.Name,
                            IsChecked = false,
                        });

                        // Process each row asynchronously
                        await Task.Run(() =>
                        {
                            for (int row = 3; row <= rowCount; row++)
                            {
                                try
                                {
                                    var numberCell = worksheet.Cells[row, 1] as Excel.Range;
                                    var dateCell = worksheet.Cells[row, 2] as Excel.Range;
                                    var documentNameCell = worksheet.Cells[row, 3] as Excel.Range;
                                    var questionCell = worksheet.Cells[row, 6] as Excel.Range;
                                    var answerCell = worksheet.Cells[row, 8] as Excel.Range;
                                    var statusCell = worksheet.Cells[row, 10] as Excel.Range;

                                    // Add the data to the collection
                                    Clarifications.Add(new Clarification
                                    {
                                        Number = numberCell != null && numberCell.Value2 != null
                                            ? int.TryParse(numberCell.Value2.ToString(), out int number) ? number : 0
                                            : 0,
                                        Date = dateCell != null && dateCell.Value2 != null
                                            ? ConvertExcelDateToDateTime(dateCell.Value2)
                                            : DateTime.MinValue,
                                        DocumentName = documentNameCell != null && documentNameCell.Value2 != null
                                            ? documentNameCell.Value2.ToString()
                                            : string.Empty,
                                        Module = worksheet.Name,
                                        Question = questionCell != null && questionCell.Value2 != null
                                            ? questionCell.Value2.ToString()
                                            : string.Empty,
                                        Answer = answerCell != null && answerCell.Value2 != null
                                            ? answerCell.Value2.ToString()
                                            : string.Empty,
                                        Status = statusCell != null && statusCell.Value2 != null
                                            ? statusCell.Value2.ToString()
                                            : string.Empty,
                                    });
                                }
                                catch (Exception ex)
                                {
                                    throw new InvalidOperationException($"Error processing row {row} in worksheet '{worksheet.Name}'.", ex);
                                }
                            }
                        });
                    }
                    else
                    {
                        MessageBox.Show(messageBoxText: $"'{worksheet.Name}' is an invalid sheet press OK to continue.",
                        caption: "Alert",
                        button: MessageBoxButton.OK,
                        icon: MessageBoxImage.Warning);
                        //throw new InvalidOperationException($"Invalid headers in sheet '{worksheet.Name}'.");
                    }
                }
            }
            catch (COMException ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw new InvalidOperationException("An error occurred while interacting with Excel. Please check the file and try again.", ex);
            }
            catch (InvalidOperationException ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw;
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw new InvalidOperationException("Access to the file is denied. Please check the file permissions and try again.", ex);
            }
            catch (Exception ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw new InvalidOperationException("An unexpected error occurred. Please try again or contact support.", ex);
            }
            finally
            {
                // Cleanup
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }

            return Clarifications;
        }

        /// <summary>
        /// Retrieves summaries of clarification data grouped by module.
        /// </summary>
        /// <returns>A collection of summary data grouped by module.</returns>
        public ObservableCollection<Summary> GetSummaries()
        {
            var summaries = from clarification in Clarifications
                    group clarification by clarification.Module into moduleGroup
                    select new Summary
                    {
                        Module = moduleGroup.Key,
                        Closed = moduleGroup.Count(c => c.Status == "Closed"),
                        Open = moduleGroup.Count(c => c.Status == "Open"),
                        OnHold = moduleGroup.Count(c => c.Status == "On Hold"),
                        Pending = moduleGroup.Count(c => c.Status == "Pending"),
                        Total = moduleGroup.Count()
                    };

            return new ObservableCollection<Summary>(summaries);
        }

        /// <summary>
        /// Retrieves the list of modules that have been processed from the Excel file.
        /// </summary>
        /// <returns>A collection of modules.</returns>
        public ObservableCollection<Models.Module> GetModules()
        {
            return Modules;
        }

        /// <summary>
        /// Converts an Excel date value to a DateTime object.
        /// </summary>
        /// <param name="excelDate">The Excel date value.</param>
        /// <returns>The converted DateTime object.</returns>
        private DateTime ConvertExcelDateToDateTime(object excelDate)
        {
            if (excelDate is double serialDate)
            {
                // Excel dates are based on the OLE Automation date
                try
                {
                    DateTime dateTime = DateTime.FromOADate(serialDate);
                    return dateTime.Date; // Return only the date part
                }
                catch
                {
                    return DateTime.MinValue; // Return a default value if conversion fails
                }
            }
            return DateTime.MinValue; // Return a default value if the input is not a valid date
        }

        /// <summary>
        /// Exports the clarification data to an Excel file.
        /// </summary>
        /// <param name="clarifications">The collection of clarifications to export.</param>
        /// <param name="filename">The file name to save the Excel document as.</param>
        public void ExportClarificationsToExcel(ObservableCollection<Clarification> clarifications, string filename)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet clarificationSheet = null;

            try
            {
                // Initialize Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = false;

                // Create a new workbook and worksheet
                workbook = excelApp.Workbooks.Add();
                clarificationSheet = (Excel.Worksheet)workbook.Sheets[1];
                clarificationSheet.Name = "Clarifications";

                // Set headers
                clarificationSheet.Cells[1, 1] = "Number";
                clarificationSheet.Cells[1, 2] = "Date";
                clarificationSheet.Cells[1, 3] = "Document Name";
                clarificationSheet.Cells[1, 4] = "Module";
                clarificationSheet.Cells[1, 5] = "Question";
                clarificationSheet.Cells[1, 6] = "Answer";
                clarificationSheet.Cells[1, 7] = "Status";

                // Fill data
                int row = 2;
                foreach (var clarification in clarifications)
                {
                    clarificationSheet.Cells[row, 1] = clarification.Number;
                    clarificationSheet.Cells[row, 2] = clarification.Date.ToShortDateString();
                    clarificationSheet.Cells[row, 3] = clarification.DocumentName;
                    clarificationSheet.Cells[row, 4] = clarification.Module;
                    clarificationSheet.Cells[row, 5] = clarification.Question;
                    clarificationSheet.Cells[row, 6] = clarification.Answer;
                    clarificationSheet.Cells[row, 7] = clarification.Status;
                    row++;
                }

                // Save the workbook
                workbook.SaveAs(filename);
            }
            catch (COMException ex)
            {
                Logger.log.Error($"{ex.Message}");
                // Handle Excel-related exceptions
                throw;
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.log.Error($"{ex.Message}");
                // Handle permission errors for file access
                throw;
            }
            catch (Exception ex)
            {
                Logger.log.Error($"{ex.Message}");
                // General exception handler for other types of errors
                throw;
            }
            finally
            {
                // Cleanup: Close and release Excel objects, even if an error occurred
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                if (clarificationSheet != null)
                {
                    Marshal.ReleaseComObject(clarificationSheet);
                }
            }
        }

        /// <summary>
        /// Exports the summary data to an Excel file.
        /// </summary>
        /// <param name="summaries">The collection of summaries to export.</param>
        /// <param name="filename">The file name to save the Excel document as.</param>
        public void ExportSummaryToExcel(ObservableCollection<Summary> summaries, string filename)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet summarySheet = null;

            try
            {
                // Initialize Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = false;

                // Create workbook and worksheet
                workbook = excelApp.Workbooks.Add();
                summarySheet = (Excel.Worksheet)workbook.Sheets[1];
                summarySheet.Name = "Summary";

                // Set headers
                summarySheet.Cells[1, 1] = "Module";
                summarySheet.Cells[1, 2] = "Closed";
                summarySheet.Cells[1, 3] = "Open";
                summarySheet.Cells[1, 4] = "On Hold";
                summarySheet.Cells[1, 5] = "Pending";
                summarySheet.Cells[1, 6] = "Total";

                // Fill data
                int row = 2;
                foreach (var summary in summaries)
                {
                    summarySheet.Cells[row, 1] = summary.Module;
                    summarySheet.Cells[row, 2] = summary.Closed;
                    summarySheet.Cells[row, 3] = summary.Open;
                    summarySheet.Cells[row, 4] = summary.OnHold;
                    summarySheet.Cells[row, 5] = summary.Pending;
                    summarySheet.Cells[row, 6] = summary.Total;
                    row++;
                }

                // Save and close workbook
                workbook.SaveAs(filename);
            }
            catch (COMException ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw; // re-throw to allow the caller to handle it
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw; // re-throw to allow the caller to handle it
            }
            catch (Exception ex)
            {
                Logger.log.Error($"{ex.Message}");
                throw; // re-throw to allow the caller to handle it
            }
            finally
            {
                // Cleanup COM objects
                if (summarySheet != null)
                {
                    Marshal.ReleaseComObject(summarySheet);
                }
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

    }
}


