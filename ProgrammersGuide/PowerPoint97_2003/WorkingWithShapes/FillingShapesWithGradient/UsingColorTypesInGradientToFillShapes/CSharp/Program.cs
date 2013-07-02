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

namespace UsingColorTypesInGradientToFillShapes
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


            //Adding an ellipse shape into the slide by defining its X,Y postion, width
            //and height
            Shape shape = slide.Shapes.AddEllipse(2300, 1200, 1000, 2000);


            //Setting the fill type of the ellipse to gradient
            shape.FillFormat.Type = FillType.Gradient;


            //Setting the color type of the gradient to two colors
            shape.FillFormat.GradientColorType = GradientColorType.TwoColors;


            //Setting the background color of the ellipse to red
            shape.FillFormat.BackColor = Color.Red;


            //Setting the foreground color of the ellipse to blue
            shape.FillFormat.ForeColor = Color.Blue;


            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");

        }
    }
}