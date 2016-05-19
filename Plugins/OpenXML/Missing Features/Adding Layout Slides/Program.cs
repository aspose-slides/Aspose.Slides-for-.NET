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
            string FilePath = @"..\..\..\Sample Files\";
            string FileName = FilePath + "Adding Layout Slides.pptx";
            
            //Instantiate Presentation class that represents the presentation file
            using (Presentation p = new Presentation(FileName))
            {
                // Try to search by layout slide type
                IMasterLayoutSlideCollection layoutSlides = p.Masters[0].LayoutSlides;
                ILayoutSlide layoutSlide =
                    layoutSlides.GetByType(SlideLayoutType.TitleAndObject) ??
                    layoutSlides.GetByType(SlideLayoutType.Title);

                if (layoutSlide == null)
                {
                    // The situation when a presentation doesn't contain some type of layouts.
                    // Technographics.pptx presentation only contains Blank and Custom layout types.
                    // But layout slides with Custom types has different slide names,
                    // like "Title", "Title and Content", etc. And it is possible to use these
                    // names for layout slide selection.
                    // Also it is possible to use the set of placeholder shape types. For example,
                    // Title slide should have only Title pleceholder type, etc.
                    foreach (ILayoutSlide titleAndObjectLayoutSlide in layoutSlides)
                    {
                        if (titleAndObjectLayoutSlide.Name == "Title and Object")
                        {
                            layoutSlide = titleAndObjectLayoutSlide;
                            break;
                        }
                    }
                    if (layoutSlide == null)
                    {
                        foreach (ILayoutSlide titleLayoutSlide in layoutSlides)
                        {
                            if (titleLayoutSlide.Name == "Title")
                            {
                                layoutSlide = titleLayoutSlide;
                                break;
                            }
                        }
                        if (layoutSlide == null)
                        {
                            layoutSlide = layoutSlides.GetByType(SlideLayoutType.Blank);
                            if (layoutSlide == null)
                            {
                                layoutSlide = layoutSlides.Add(SlideLayoutType.TitleAndObject, "Title and Object");
                            }
                        }
                    }
                }

                //Adding empty slide with added layout slide 
                p.Slides.InsertEmptySlide(0, layoutSlide);

                //Save presentation    
                p.Save(FileName, SaveFormat.Pptx);
            }
        }
    }
}
