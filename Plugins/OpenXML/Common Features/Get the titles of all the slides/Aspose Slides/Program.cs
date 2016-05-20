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
            string FileName = FilePath + "Get the titles of all the slides.pptx";
            foreach (string s in GetSlideTitles(FileName))
                Console.WriteLine(s);
            Console.ReadKey();
        }
        // Get a list of the titles of all the slides in the presentation.
        public static IList<string> GetSlideTitles(string presentationFile)
        {
            // Create a new linked list of strings.
            List<string> texts = new List<string>();

            //Instantiate PresentationEx class that represents PPTX
            using (Presentation pres = new Presentation(presentationFile))
            {

                //Access all the slides
                foreach (ISlide sld in pres.Slides)
                {

                    //Iterate through shapes to find the placeholder
                    foreach (Shape shp in sld.Shapes)
                        if (shp.Placeholder != null)
                        {
                            if (IsTitleShape(shp))
                            {
                                //get the text of placeholder
                                texts.Add(((AutoShape)shp).TextFrame.Text);
                            }
                        }
                }
            }

            // Return an array of strings.
            return texts;
        }
        // Determines whether the shape is a title shape.
        private static bool IsTitleShape(Shape shape)
        {
            switch (shape.Placeholder.Type)
            {
                case PlaceholderType.Title:
                case PlaceholderType.CenteredTitle:
                    return true;
                default:
                    return false;
            }
        }
    }
}
