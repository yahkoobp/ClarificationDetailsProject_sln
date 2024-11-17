using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPTMerger.Repo;

namespace PPTMerger.MergeRepo
{
    internal class PPTMergeRepo : IRepo
    {
        public void MergeFiles(ObservableCollection<string> pptPaths, string outputPath)
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation mergedPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            try
            {
                foreach (string pptPath in pptPaths)
                {
                    //open each presentations in the path
                    Presentation sourcePresentation = pptApplication.Presentations.Open(
                        pptPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    //for each slides in the presentaion do the following
                    foreach (Slide slide in sourcePresentation.Slides)
                    {
                        //copy the slide
                        slide.Copy();
                        //Paste the slide to the merged presentation
                        Slide newSlide = mergedPresentation.Slides.Paste(mergedPresentation.Slides.Count + 1)[1];
                        //Preserves the source format and designs
                        newSlide.Design = slide.Design;
                        // newSlide.CustomLayout = slide.CustomLayout;
                        //mergedPresentation.Slides.Paste(mergedPresentation.Slides.Count + 1);
                    }

                    sourcePresentation.Close();
                    Marshal.ReleaseComObject(sourcePresentation);
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
    }
}
