using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Export;
namespace ConsoleApplication19
{
    class SetSlideSizeScale
    {
        static void Main(string[] args)
        {

            //ExStart:SetSlideSizeScale
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Layout();

            // ExStart:SettSizeAndType
            // Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(dataDir + "AccessSlides.pptx");
            Presentation auxPresentation = new Presentation();

            ISlide slide = presentation.Slides[0];

            // Set the slide size of generated presentations to that of source
            presentation.SlideSize.SetSize(540, 720, SlideSizeScaleType.EnsureFit); // Method SetSize is used for set slide size with scale content to ensure fit
            presentation.SlideSize.SetSize(SlideSizeType.A4Paper, SlideSizeScaleType.Maximize); // Method SetSize is used for set slide size with maximize size of content

          
           
            // Save Presentation to disk
            auxPresentation.Save(dataDir + "Set_Size&Type_out.pptx", SaveFormat.Pptx);
            //ExEnd:SetSlideSizeScale
            
            
        
            
           
        
        }
    }
}
