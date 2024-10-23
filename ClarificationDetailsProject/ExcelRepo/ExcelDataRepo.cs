// ExcelRepo.cs
// 
// This file contains the implementation of the ExcelRepo class,
// which is responsible for handling data operations related to
// Excel files. It implements the IRepo interface and provides
// methods for loading, filtering, and searching data within
// Excel documents.
// 
// Author: Yahkoob P
// Date: 2024-10-23

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ClarificationDetailsProject.Models;
using ClarificationDetailsProject.Repo;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
        ObservableCollection<Clarification> Clarifications = new ObservableCollection<Clarification>();
        private List<string> expectedHeaders = new List<string>() 
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

        public void ApplyFilters()
        {
            throw new NotImplementedException();
        }
       
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
            catch (Exception)
            {
                throw;
            }
        }

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

        public ObservableCollection<Clarification> LoadData(string filePath)
        {
            ObservableCollection<Clarification> clarifications = new ObservableCollection<Clarification>();
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                // Open the workbook
                workbook = excelApp.Workbooks.Open(filePath);

                // Check if the workbook has any worksheets
                if (workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("The Excel file does not contain any worksheets.");
                }

                // Loop through each worksheet in the workbook
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    // Validate the worksheet headers
                    if (IsValidExcelWorkBook(worksheet))
                    {
                        Excel.Range usedRange = worksheet.UsedRange;
                        if (usedRange == null || usedRange.Rows.Count == 0)
                        {
                            throw new InvalidOperationException($"The worksheet '{worksheet.Name}' does not contain any data.");
                        }

                        int rowCount = usedRange.Rows.Count;

                        // Loop through the rows starting from row 3
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
                                clarifications.Add(new Clarification
                                {
                                    Number = numberCell != null && numberCell.Value2 != null ?
                                        int.TryParse(numberCell.Value2.ToString(), out int number) ? number : 0 : 0,
                                    Date = dateCell != null && dateCell.Value2 != null ?
                                        DateTime.TryParse(dateCell.Value2.ToString(), out DateTime date) ? date : DateTime.MinValue : DateTime.MinValue,
                                    DocumentName = documentNameCell != null && documentNameCell.Value2 != null ?
                                        documentNameCell.Value2.ToString() : string.Empty,
                                    Module = (worksheet.Cells[1, 1] as Excel.Range).Value2?.ToString() ?? string.Empty,
                                    Question = questionCell != null && questionCell.Value2 != null ?
                                        questionCell.Value2.ToString() : string.Empty,
                                    Answer = answerCell != null && answerCell.Value2 != null ?
                                        answerCell.Value2.ToString() : string.Empty,
                                    Status = statusCell != null && statusCell.Value2 != null ?
                                        statusCell.Value2.ToString() : string.Empty,
                                });
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing row {row} in worksheet '{worksheet.Name}': {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Invalid headers in sheet '{worksheet.Name}'. Press OK to continue.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading data from Excel: {ex.Message}");
            }
            finally
            {
                // Cleanup
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }

            return clarifications;
        }

        public void Search(string text)
        {
            throw new NotImplementedException();
        }
    }
}
