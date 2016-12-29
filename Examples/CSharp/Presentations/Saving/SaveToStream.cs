using System.IO;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Saving
{
    class SaveToStream
    {
        public static void Run()
        {
            //ExStart:SaveToStream
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();

            // Instantiate a Presentation object that represents a PPT file
            using (Presentation presentation = new Presentation())
            {

                IAutoShape shape = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 200, 200);

                // Add text to shape
                shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream.";

                FileStream toStream = new FileStream(dataDir + "Save_As_Stream_out.pptx", FileMode.Create);
                presentation.Save(toStream, Aspose.Slides.Export.SaveFormat.Pptx);
                toStream.Close();
            }
            //ExEnd:SaveToStream
        }
    }
}
