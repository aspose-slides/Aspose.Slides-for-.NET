using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string FileName = FilePath + "Managing Slides Transitions.pptx";
            
            //Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(FileName))
            {

                //Apply circle type transition on slide 1
                pres.Slides[0].SlideShowTransition.Type = TransitionType.Circle;

                //Apply comb type transition on slide 2
                pres.Slides[1].SlideShowTransition.Type = TransitionType.Comb;

                //Apply zoom type transition on slide 3
                pres.Slides[2].SlideShowTransition.Type = TransitionType.Zoom;

                //Write the presentation to disk
                pres.Save(FileName, SaveFormat.Pptx);

            }
        }
    }
}
