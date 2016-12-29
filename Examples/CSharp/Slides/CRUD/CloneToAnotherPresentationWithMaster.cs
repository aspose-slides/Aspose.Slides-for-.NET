using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/


namespace Aspose.Slides.Examples.CSharp.Slides.CRUD
{
    public class CloneToAnotherPresentationWithMaster
    {
        public static void Run()
        {
            //ExStart:CloneToAnotherPresentationWithMaster
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // Instantiate Presentation class to load the source presentation file

            using (Presentation srcPres = new Presentation(dataDir + "CloneToAnotherPresentationWithMaster.pptx"))
            {
                // Instantiate Presentation class for destination presentation (where slide is to be cloned)
                using (Presentation destPres = new Presentation())
                {

                    // Instantiate ISlide from the collection of slides in source presentation along with
                    // Master slide
                    ISlide SourceSlide = srcPres.Slides[0];
                    IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

                    // Clone the desired master slide from the source presentation to the collection of masters in the
                    // Destination presentation
                    IMasterSlideCollection masters = destPres.Masters;
                    IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

                    // Clone the desired master slide from the source presentation to the collection of masters in the
                    // Destination presentation
                    IMasterSlide iSlide = masters.AddClone(SourceMaster);

                    // Clone the desired slide from the source presentation with the desired master to the end of the
                    // Collection of slides in the destination presentation
                    ISlideCollection slds = destPres.Slides;
                    slds.AddClone(SourceSlide, iSlide, true);
                  
                    // Clone the desired master slide from the source presentation to the collection of masters in the // Destination presentation
                    // Save the destination presentation to disk
                    destPres.Save(dataDir + "CloneToAnotherPresentationWithMaster_out.pptx", SaveFormat.Pptx);

                }
            }
            //ExEnd:CloneToAnotherPresentationWithMaster
        }
    }
}