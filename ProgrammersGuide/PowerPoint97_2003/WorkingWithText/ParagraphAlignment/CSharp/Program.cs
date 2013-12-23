//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ParagraphAlignment
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
            
            //Aligning the text paragraph to center
            para.Alignment = TextAlignment.Center;
            
            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");

        }
    }
}