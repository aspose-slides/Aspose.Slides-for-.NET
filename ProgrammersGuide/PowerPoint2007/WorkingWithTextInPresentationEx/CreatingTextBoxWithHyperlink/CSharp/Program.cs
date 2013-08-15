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

namespace CreatingTextBoxWithHyperlink
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

            //Instantiate a PresentationEx class that represents a PPTX
            PresentationEx pptxPresentation = new PresentationEx();

            //Get first slide
            SlideEx slide = pptxPresentation.Slides[0];

            //Add an AutoShape of Rectangle Type
            int index = slide.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 150, 150, 50);

            //Get a reference to the added shape 
            Aspose.Slides.Pptx.ShapeEx pptxShape = slide.Shapes[index];

            //Cast the shape to AutoShape
            AutoShapeEx pptxAutoShape = (AutoShapeEx)pptxShape;

            //Access TextFrame associated with the AutoShape
            TextFrameEx TextFrame = pptxAutoShape.TextFrame;

            //Add some text to the frame
            TextFrame.Paragraphs[0].Portions[0].Text = "Aspose.Slides";

            //Set Hyperlink for the portion text
            HyperlinkEx HLink = new HyperlinkEx("http://www.aspose.com");
            TextFrame.Paragraphs[0].Portions[0].HLinkClick = HLink;

            //Save the PPTX Presentation
            pptxPresentation.Write(dataDir + "output.pptx");

        }
    }
}