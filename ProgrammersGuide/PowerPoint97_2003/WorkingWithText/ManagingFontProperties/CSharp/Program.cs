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

namespace ManagingFontProperties
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
            Slide slide = pres.GetSlideByPosition(1);


            //Accessing the first placeholder in the slide and typecasting it as a text holder
            TextHolder th = (TextHolder)slide.Placeholders[0];


            //Getting the first paragraph of the text holder
            Paragraph para = th.Paragraphs[0];


            //Accessing the first portion of the paragraph
            Portion port = para.Portions[0];


            //Setting the font to bold
            port.FontBold = true;


            //Setting the font to be underlined
            port.FontUnderline = true;


            //Setting the font to be embossed
            port.FontEmbossed = true;


            //Setting the font to be italic
            port.FontItalic = true;


            //Enabling the font shadow
            port.FontShadow = true;


            //Setting the font color to green
            port.FontColor = Color.Green;


            //Setting the font height to 55
            port.FontHeight = 55;


            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");

        }
    }
}