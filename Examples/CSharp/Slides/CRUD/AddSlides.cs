using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.CRUD
{
    public class AddSlides
    {
        public static void Run()
        {
            //ExStart:AddSlides
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Presentation class that represents the presentation file
            using (Presentation pres = new Presentation())
            {
                // Instantiate SlideCollection calss
                ISlideCollection slds = pres.Slides;

                for (int i = 0; i < pres.LayoutSlides.Count; i++)
                {
                    // Add an empty slide to the Slides collection
                    slds.AddEmptySlide(pres.LayoutSlides[i]);

                }

                // Save the PPTX file to the Disk
                pres.Save(dataDir + "EmptySlide_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
            //ExEnd:AddSlides
        }
    }
}