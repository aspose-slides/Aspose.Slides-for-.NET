// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using System;
using System.Collections.Generic;

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
            foreach (string s in GetAllTextInSlide(FileName, 0))
                Console.WriteLine(s);
            Console.ReadKey();
        }
        // Get all the text in a slide.
        public static List<string> GetAllTextInSlide(string presentationFile, int slideIndex)
        {
            // Create a new linked list of strings.
            List<string> texts = new List<string>();

            //Instantiate PresentationEx class that represents PPTX
            using (Presentation pres = new Presentation(presentationFile))
            {

                //Access the slide
                ISlide sld = pres.Slides[slideIndex];

                //Iterate through shapes to find the placeholder
                foreach (Shape shp in sld.Shapes)
                    if (shp.Placeholder != null)
                    {
                        //get the text of each placeholder
                        texts.Add(((AutoShape)shp).TextFrame.Text);
                    }

            }

            // Return an array of strings.
            return texts;
        }
    }
}
