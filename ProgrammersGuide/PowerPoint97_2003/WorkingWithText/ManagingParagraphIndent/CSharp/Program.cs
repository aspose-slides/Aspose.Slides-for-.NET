//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ManagingParagraphIndent
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

            //Instantiate Presentation Class
            Presentation pres = new Presentation();

            //Get first slide
            Slide sld = pres.GetSlideByPosition(1);

            //Add a Rectangle Shape
            Aspose.Slides.Rectangle rect = sld.Shapes.AddRectangle(500, 500, 1500, 150);

            //Add TextFrame to the Rectangle
            TextFrame tf = rect.AddTextFrame("This is first line \r This is second line \r This is third line");

            //Set the text to fit the shape
            tf.FitShapeToText = true;

            //Hide the lines of the Rectangle
            tf.LineFormat.ShowLines = false;

            //Get first Paragraph in the TextFrame and set its Indent
            Paragraph para1 = tf.Paragraphs[0];
            para1.BulletOffset = 150;

            //Get second Paragraph in the TextFrame and set its Indent
            Paragraph para2 = tf.Paragraphs[1];
            para2.BulletOffset = 250;

            //Get third Paragraph in the TextFrame and set its Indent 
            Paragraph para3 = tf.Paragraphs[2];
            para3.BulletOffset = 350;

            //Write the Presentation to disk
            pres.Write(dataDir + "output.ppt");

        }
    }
}