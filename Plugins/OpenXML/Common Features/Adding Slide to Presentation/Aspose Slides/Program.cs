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
            string FileName = FilePath + "Adding Slide to Presentation.pptx";
            
            //Instantiate PresentationEx class that represents the PPT file
            Presentation pres = new Presentation();

            //Blank slide is added by default, when you create
            //presentation from default constructor
            //Adding an empty slide to the presentation and getting the reference of
            //that empty slide
            ISlide slide = pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);

            

            //Write the output to disk
            pres.Save(FileName,Aspose.Slides.Export.SaveFormat.Pptx);

        }
    }
}