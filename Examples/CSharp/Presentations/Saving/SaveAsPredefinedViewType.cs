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
    class SaveAsPredefinedViewType
    {
        public static void Run()
        {
            //ExStart:SaveAsPredefinedViewType
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();

            // Opening the presentation file
            Presentation presentation = new Presentation();

            // Setting view type 
            presentation.ViewProperties.LastView = ViewType.SlideMasterView;

            // Saving presentation
            presentation.Save(dataDir + "SetViewType_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            //ExEnd:SaveAsPredefinedViewType
        }
    }
}
