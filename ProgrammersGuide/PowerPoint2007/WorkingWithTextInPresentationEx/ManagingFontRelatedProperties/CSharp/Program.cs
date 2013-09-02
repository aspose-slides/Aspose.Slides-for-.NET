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
using System.Drawing;

namespace ManagingFontRelatedProperties
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");


            //Accessing a slide using its slide position
            SlideEx slide = pres.Slides[0];

            //Accessing the first and second placeholder in the slide and typecasting it as AutoShape
            TextFrameEx tf1 = ((AutoShapeEx)slide.Shapes[0]).TextFrame;
            TextFrameEx tf2 = ((AutoShapeEx)slide.Shapes[1]).TextFrame;

            //Accessing the first Paragraph
            ParagraphEx para1 = tf1.Paragraphs[0];
            ParagraphEx para2 = tf2.Paragraphs[0];

            //Accessing the first portion
            PortionEx port1 = para1.Portions[0];
            PortionEx port2 = para2.Portions[0];

            //Define new fonts
            FontDataEx fd1 = new FontDataEx("Elephant");
            FontDataEx fd2 = new FontDataEx("Castellar");

            //Assign new fonts to portion
            port1.PortionFormat.LatinFont = fd1;
            port2.PortionFormat.LatinFont = fd2;

            //Set font to Bold
            port1.PortionFormat.FontBold = NullableBool.True;
            port2.PortionFormat.FontBold = NullableBool.True;

            //Set font to Italic
            port1.PortionFormat.FontItalic = NullableBool.True;
            port2.PortionFormat.FontItalic = NullableBool.True;

            //Set font color
            port1.PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
            port1.PortionFormat.FillFormat.SolidFillColor.Color = Color.Purple;
            port2.PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
            port2.PortionFormat.FillFormat.SolidFillColor.Color = Color.Peru;

            //Write the PPTX to disk
            pres.Write(dataDir + "output.pptx");


        }
    }
}