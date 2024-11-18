using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPTMerger.Repo;

namespace PPTMerger.MergeRepo
{
    internal class PPTMergeRepo : IRepo
    {
        public Delegate RaiseError;

        public void MergeFiles(ObservableCollection<string> pptPaths, string outputPath)
        {          
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation mergedPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            try
            {
                foreach (string pptPath in pptPaths)
                {
                    try
                    {
                        if (!IsPowerpointFile(pptPath))
                        {
                            MessageBox.Show($"{pptPath} is not a valid powerpoint file , press ok to continue");
                            continue;
                        }
                        //open each presentations in the path
                        Presentation sourcePresentation = pptApplication.Presentations.Open(
                            pptPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                        //for each slides in the presentaion do the following
                        foreach (Slide slide in sourcePresentation.Slides)
                        {
                            try
                            {
                                //copy the slide
                                slide.Copy();
                                //Paste the slide to the merged presentation
                                Slide newSlide = mergedPresentation.Slides.Paste(mergedPresentation.Slides.Count + 1)[1];
                                //Preserves the source format and designs
                                newSlide.Design = slide.Design;
                            }
                            catch(Exception ex)
                            {
                                continue;
                            }
                           
                        }
                        sourcePresentation.Close();
                        Marshal.ReleaseComObject(sourcePresentation);
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                    
                }
                //Save the merged presentation
                mergedPresentation.SaveAs(outputPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //Cleanup the resources
                mergedPresentation.Close();
                pptApplication.Quit();
                Marshal.ReleaseComObject(mergedPresentation);
                Marshal.ReleaseComObject(pptApplication);
            }
        }

        public bool IsPowerpointFile(string filepath)
        {
            string extension = Path.GetExtension(filepath)?.ToLower();
            return extension == ".ppt" || extension == ".pptx" || extension == ".pps" || extension == ".ppsx";
        }
    }
}
