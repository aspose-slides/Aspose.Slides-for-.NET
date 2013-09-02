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

namespace ManagingParagraphsAlignment
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");


            //Accessing first slide
            SlideEx slide = pres.Slides[0];

            //Accessing the first and second placeholder in the slide and typecasting it as AutoShape
            TextFrameEx tf1 = ((AutoShapeEx)slide.Shapes[0]).TextFrame;
            TextFrameEx tf2 = ((AutoShapeEx)slide.Shapes[1]).TextFrame;

            //Change the text in both placeholders
            tf1.Text = "Center Align by Aspose";
            tf2.Text = "Center Align by Aspose";

            //Getting the first paragraph of the placeholders
            ParagraphEx para1 = tf1.Paragraphs[0];
            ParagraphEx para2 = tf2.Paragraphs[0];

            //Aligning the text paragraph to center
            para1.ParagraphFormat.Alignment = TextAlignmentEx.Center;
            para2.ParagraphFormat.Alignment = TextAlignmentEx.Center;

            //Writing the presentation as a PPTX file
            pres.Write(dataDir + "output.pptx");


        }
    }
}