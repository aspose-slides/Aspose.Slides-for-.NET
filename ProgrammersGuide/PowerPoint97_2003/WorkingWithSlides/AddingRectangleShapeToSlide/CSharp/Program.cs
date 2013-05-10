//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing;

namespace AddingRectangleShapeToSlide
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");
            
            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(2);

            //Adding a rectangle shape into the slide by defining its X,Y postion, width and height
            slide.Shapes.AddRectangle(1400, 600, 3000, 1000);

            //Adding a rectangle shape into the slide by defining its X,Y postion, width and height
            Shape shape = slide.Shapes.AddRectangle(1400, 1300, 3000, 1000);

            //Setting the fill type of the rectangle to pattern
            shape.FillFormat.Type = FillType.Pattern;
            
            //Setting the pattern style to sphere
            shape.FillFormat.PatternStyle = PatternStyle.Sphere;

            //Setting the background color of the rectangle to light gray
            shape.FillFormat.BackColor = Color.LightGray;

            //Setting the foreground color of the rectangle to yellow
            shape.FillFormat.ForeColor = Color.Yellow;

            //Setting the foreground color of the rectangle lines to blue
            shape.LineFormat.ForeColor = Color.Blue;

            //Setting the width of the rectangle lines in points
            shape.LineFormat.Width = 10;

            //Setting the line style of the rectangle to thick thin
            shape.LineFormat.Style = LineStyle.ThickThin;

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");
            
        }
    }
}