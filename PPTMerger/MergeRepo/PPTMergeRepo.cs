using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPTMerger.Delegates;
using PPTMerger.Repo;

namespace PPTMerger.MergeRepo
{
    internal class PPTMergeRepo : IRepo
    {
        public event EventHandler<FileProcessingFailedEventArgs> FileProcessingFailed;
        public event Action<string> LogEvent;
        public event EventHandler<int> ProgressEvent;
        protected void OnLog(string message)
        {
            LogEvent?.Invoke($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        }
        public async Task MergeFilesAsync(ObservableCollection<string> pptPaths, string outputPath)
        {
            await Task.Run(() =>
            {
                int totalFiles = pptPaths.Count;
                int currentFile = 0;
                Microsoft.Office.Interop.PowerPoint.Application pptApplication = null;
                Presentation mergedPresentation = null;

                try
                {
                    pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
                    mergedPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);

                    OnLog($"Starting merge operation with {pptPaths.Count} files.");
                    int processedCount = 0;
                    int skippedCount = 0;

                    foreach (string pptPath in pptPaths)
                    {
                        try
                        {
                            currentFile++;
                            if (!IsPowerpointFile(pptPath))
                            {
                                OnLog($"{pptPath} is not a valid PowerPoint file.");
                                skippedCount++;
                                continue;
                            }

                            // Open each presentation in the path
                            Presentation sourcePresentation = pptApplication.Presentations.Open(
                                pptPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                            OnLog($"Processing file: {pptPath}");

                            // Process each slide in the presentation
                            foreach (Slide slide in sourcePresentation.Slides)
                            {
                                try
                                {
                                    slide.Copy();
                                    Slide newSlide = mergedPresentation.Slides.Paste(mergedPresentation.Slides.Count + 1)[1];
                                    newSlide.Design = slide.Design;
                                }
                                catch
                                {
                                    OnLog($"Failed to process slide: {slide.SlideNumber}");
                                }
                            }

                            sourcePresentation.Close();
                            Marshal.ReleaseComObject(sourcePresentation);
                            processedCount++;
                        }
                        catch (Exception ex)
                        {
                            OnLog($"Error processing file: {pptPath}. Skipping. Details: {ex.Message}");
                            skippedCount++;
                        }
                        int progress = (currentFile * 100) / totalFiles;
                        OnLog($"Processing file {currentFile} of {totalFiles}...");
                        ProgressEvent?.Invoke(this, progress);
                    }

                    // Save the merged presentation
                    if (mergedPresentation.Slides.Count == 0)
                    {
                        OnLog("Merge Failed: No slides merged.");
                        throw new Exception("Merge Failed: No slides merged.");
                    }

                    mergedPresentation.SaveAs(outputPath);
                    OnLog($"Merge completed successfully. Files processed: {processedCount}, Skipped: {skippedCount}");
                }
                catch (Exception ex)
                {
                    OnLog($"Merge failed. Error: {ex.Message}");
                    throw;
                }
                finally
                {
                    // Cleanup resources
                    mergedPresentation?.Close();
                    pptApplication?.Quit();

                    if (mergedPresentation != null)
                        Marshal.ReleaseComObject(mergedPresentation);

                    if (pptApplication != null)
                        Marshal.ReleaseComObject(pptApplication);
                }
            });
        }



        public bool IsPowerpointFile(string filepath)
        {
            string extension = Path.GetExtension(filepath)?.ToLower();
            return extension == ".ppt" || extension == ".pptx" || extension == ".pps" || extension == ".ppsx";
        }
    }
}
