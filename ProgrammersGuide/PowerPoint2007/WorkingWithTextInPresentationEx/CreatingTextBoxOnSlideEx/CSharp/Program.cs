//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;

namespace CreatingTextBoxOnSlideEx
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

            //Instantiate PresentationEx
            PresentationEx pres = new PresentationEx();

            //Get the first slide
            SlideEx sld = pres.Slides[0];

            //Add an AutoShape of Rectangle type
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 75, 150, 50);
            AutoShapeEx ashp = (AutoShapeEx)sld.Shapes[idx];

            //Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");

            //Accessing the text frame
            TextFrameEx txtFrame = ashp.TextFrame;

            //Create the Paragraph object for text frame
            ParagraphEx para = txtFrame.Paragraphs[0];

            //Create Portion object for paragraph
            PortionEx portion = para.Portions[0];

            //Set Text
            portion.Text = "Aspose TextBox";
            
            //Write the presentation to disk
            pres.Write(dataDir + "output.pptx");

        }
    }
}