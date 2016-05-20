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
            string FileName = FilePath + "Delete a slide.pptx";
            DeleteSlide(FileName, 1);
        }
        
        public static void DeleteSlide(string presentationFile, int slideIndex)
        {
            //Instantiate a PresentationEx object that represents a PPTX file
            using (Presentation pres = new Presentation(presentationFile))
            {

                //Accessing a slide using its index in the slides collection
                ISlide slide = pres.Slides[slideIndex];


                //Removing a slide using its reference
                pres.Slides.Remove(slide);


                //Writing the presentation as a PPTX file
                pres.Save(presentationFile,Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}
