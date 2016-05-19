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
            string FileName = FilePath + "Get all the text in a slide.pptx";

            int numberOfSlides = CountSlides(FileName);
            System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
            string slideText;
            for (int i = 0; i < numberOfSlides; i++)
            {
                slideText = GetSlideText(FileName, i);
                System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);
            }
            System.Console.ReadKey();
        }
        public static int CountSlides(string presentationFile)
        {
            //Instantiate PresentationEx class that represents PPTX
            using (Presentation pres = new Presentation(presentationFile))
            {
                return pres.Slides.Count;
            }
        }
        public static string GetSlideText(string docName, int index)
        {
            string sldText = "";
            //Instantiate PresentationEx class that represents PPTX
            using (Presentation pres = new Presentation(docName))
            {
                //Access the slide
                ISlide sld = pres.Slides[index];

                //Iterate through shapes to find the placeholder
                foreach (Shape shp in sld.Shapes)
                    if (shp.Placeholder != null)
                    {
                        //get the text of each placeholder
                        sldText += ((AutoShape)shp).TextFrame.Text;
                    }

            }
            return sldText;
        }
    }
}
