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
            string FileName = FilePath + "Move a slide to a new position.pptx";
            MoveSlide(FileName, 1, 2);
        }
        // Move a slide to a different position in the slide order in the presentation.
        public static void MoveSlide(string presentationFile, int from, int to)
        {
            //Instantiate PresentationEx class to load the source PPTX file
            using (Presentation pres = new Presentation(presentationFile))
            {
                //Get the slide whose position is to be changed
                ISlide sld = pres.Slides[from];
                ISlide sld2 = pres.Slides[to];

                //Set the new position for the slide
                sld2.SlideNumber = from;
                sld.SlideNumber = to;

                //Write the PPTX to disk
                pres.Save(presentationFile,Aspose.Slides.Export.SaveFormat.Pptx);

            }
        }
    }
}
