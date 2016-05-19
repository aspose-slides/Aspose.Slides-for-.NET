using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Apply Theme to Presentation.pptx";
            string ThemeFileName = FilePath + "Theme.pptx";

            ApplyThemeToPresentation(ThemeFileName, FileName);
        }

        public static void ApplyThemeToPresentation(string presentationFile, string outputFile)
        {
            //Instantiate Presentation class to load the source presentation file
            Presentation srcPres = new Presentation(presentationFile);

            //Instantiate Presentation class for destination presentation (where slide is to be cloned)
            Presentation destPres = new Presentation(outputFile);

            //Instantiate ISlide from the collection of slides in source presentation along with
            //master slide
            ISlide SourceSlide = srcPres.Slides[0];

            //Clone the desired master slide from the source presentation to the collection of masters in the
            //destination presentation
            IMasterSlideCollection masters = destPres.Masters;
            IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

            //Clone the desired master slide from the source presentation to the collection of masters in the
            //destination presentation
            IMasterSlide iSlide = masters.AddClone(SourceMaster);

            //Clone the desired slide from the source presentation with the desired master to the end of the
            //collection of slides in the destination presentation
            ISlideCollection slds = destPres.Slides;
            slds.AddClone(SourceSlide, iSlide, true);

            //Clone the desired master slide from the source presentation to the collection of masters in the//destination presentation
            //Save the destination presentation to disk
            destPres.Save(outputFile, SaveFormat.Pptx);
        }
    }
}
