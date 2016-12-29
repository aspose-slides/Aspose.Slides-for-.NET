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
    public class SaveAsReadOnly
    {
        public static void Run()
        {
            //ExStart:SaveAsReadOnly
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate a Presentation object that represents a PPT file
            Presentation presentation = new Presentation();

            //....do some work here.....

            // Setting Write protection Password
            presentation.ProtectionManager.SetWriteProtection("test");

            // Save your presentation to a file
            presentation.Save(dataDir + "WriteProtected_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            //ExEnd:SaveAsReadOnly
        }
    }
}