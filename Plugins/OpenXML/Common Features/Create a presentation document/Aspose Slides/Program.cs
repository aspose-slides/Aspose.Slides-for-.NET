// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;

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
            string FileName = FilePath + "Create a presentation document.pptx";
            CreatePresentation(FileName);
        }
        public static void CreatePresentation(string filepath)
        {
            //Instantiate a Presentation object that represents a PPT file
            using (Presentation pres = new Presentation())
            {
                //Instantiate SlideExCollection calss
                ISlideCollection slds = pres.Slides;

                //Add an empty slide to the SlidesEx collection
                slds.AddEmptySlide(pres.LayoutSlides[0]);

                //Save your presentation to a file
                pres.Save(filepath,Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}
