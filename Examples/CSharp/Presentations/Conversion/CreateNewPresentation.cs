using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class CreateNewPresentation
    {
        public static void Run()
        {
            //ExStart:ConvetToSWF
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation())
            {
                // Get the first slide
                ISlide slide = presentation.Slides[0];

                // Add an autoshape of type line
                slide.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0);
                presentation.Save(dataDir + "NewPresentation_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:ConvetToSWF
        } 
    }
}

 