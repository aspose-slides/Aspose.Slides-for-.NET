using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Adding_Layout_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate Presentation class that represents the presentation file
            using (Presentation p = new Presentation("Test.pptx"))
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
                p.Save("Output.pptx", SaveFormat.Pptx);
            }
        }
    }
}
