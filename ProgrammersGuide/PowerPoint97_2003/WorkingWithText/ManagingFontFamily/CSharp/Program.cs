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

namespace ManagingFontFamily
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

            //Instantiate Presentation class
            Presentation pres = new Presentation();

            //Get first slide
            Slide sld = pres.GetSlideByPosition(1);

            //Add a Rectangle shape to the slide
            Aspose.Slides.Rectangle rect = sld.Shapes.AddRectangle(500, 500, 1500, 75);

            //Add a TextFrame to the Rectangle
            TextFrame tf = rect.AddTextFrame("Aspose Text Box");

            //Resize the Rectangle to fit text
            tf.FitShapeToText = true;

            //Get the Portion object associated with the TextFrame
            Portion port = tf.Paragraphs[0].Portions[0];

            //Set the font of the portion
            pres.Fonts[port.FontIndex].FontName = "Times New Roman";

            //Set Bold property of the Font
            port.FontBold = true;

            //Set Italic property of the Font
            port.FontItalic = true;

            //Set Underline property of the Font
            port.FontUnderline = true;

            //Set Height of the Font
            port.FontHeight = 25;

            //Set the Color of the Font
            port.FontColor = Color.Blue;

            //Write the Presentation to the disk
            pres.Write(dataDir + "output.ppt");

        }
    }
}