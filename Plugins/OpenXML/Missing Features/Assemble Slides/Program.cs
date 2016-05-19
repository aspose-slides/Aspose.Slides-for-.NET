// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        private static string MyDir = @"..\..\..\Sample Files\";

        static void Main(string[] args)
        {
            AddingSlidetoPresentation();
            AccessingSlidesOfPresentation();
            RemovingSlides();
            ChangingPositionOfSlide();
        }
        public static void AddingSlidetoPresentation()
        {
            Presentation pres = new Presentation();

            //Instantiate SlideCollection class

            ISlideCollection slds = pres.Slides;
            for (int i = 0; i < pres.LayoutSlides.Count; i++)
            {
                //Add an empty slide to the Slides collection
                slds.AddEmptySlide(pres.LayoutSlides[i]);

            }

            //Save the PPTX file to the Disk
            pres.Save(MyDir + "Assemble Slides.pptx", SaveFormat.Pptx);
        }
        public static void AccessingSlidesOfPresentation()
        {
            //Instantiate a Presentation object that represents a presentation file
            Presentation pres = new Presentation(MyDir + "Assemble Slides.pptx");
            //Accessing a slide using its slide index
            ISlide slide = pres.Slides[0];

        }
        public static void RemovingSlides()
        {
            //Instantiate a Presentation object that represents a presentation file
            Presentation pres = new Presentation(MyDir + "Assemble Slides.pptx");

            //Accessing a slide using its index in the slides collection
            ISlide slide = pres.Slides[0];
            //Removing a slide using its reference
            pres.Slides.Remove(slide);

            //Writing the presentation file
            pres.Save(MyDir + "Assemble Slides.pptx", SaveFormat.Pptx);

        }
        public static void ChangingPositionOfSlide()
        {
            //Instantiate Presentation class to load the source presentation file
            Presentation pres = new Presentation(MyDir + "Assemble Slides.pptx");
            {
                //Get the slide whose position is to be changed
                ISlide sld = pres.Slides[0];
                //Set the new position for the slide
                sld.SlideNumber = 2;
                //Write the presentation to disk
                pres.Save(MyDir + "Assemble Slides.pptx", SaveFormat.Pptx);
            }

        }
    }
}
