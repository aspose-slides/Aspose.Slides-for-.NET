//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace SavePresentationToFile
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);
				
            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation();

            //....do some work here.....
            //Adding an empty slide to the presentation and getting the reference of
            //that empty slide
            Slide slide = pres.AddEmptySlide();
            //Adding a rectangle (X=2400, Y=1800, Width=1000 & Height=500) to the slide
            Aspose.Slides.Rectangle rect = slide.Shapes.AddRectangle(2400, 1800, 1000, 500);
            //Hiding the lines of rectangle
            rect.LineFormat.ShowLines = false;
            //Adding a text frame to the rectangle with "Hello World" as a default text
            rect.AddTextFrame("Hello World");
            //Removing the first slide of the presentation which is always added by
            //Aspose.Slides for .NET by default while creating the presentation
            pres.Slides.RemoveAt(0);

            //Save your presentation to a file
            pres.Write(dataDir + "demo.ppt");
        }
    }
}