using System.Drawing;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Background
{
    public class SetSlideBackgroundMaster
    {
        public static void Run()
        {
            //ExStart:SetSlideBackgroundMaster
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Background();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate the Presentation class that represents the presentation file
            using (Presentation pres = new Presentation())
            {

                // Set the background color of the Master ISlide to Forest Green
                pres.Masters[0].Background.Type = BackgroundType.OwnBackground;
                pres.Masters[0].Background.FillFormat.FillType = FillType.Solid;
                pres.Masters[0].Background.FillFormat.SolidFillColor.Color = Color.ForestGreen;

                // Write the presentation to disk
                pres.Save(dataDir + "SetSlideBackgroundMaster_out.pptx", SaveFormat.Pptx);

            }
            //ExEnd:SetSlideBackgroundMaster
        }
    }
}