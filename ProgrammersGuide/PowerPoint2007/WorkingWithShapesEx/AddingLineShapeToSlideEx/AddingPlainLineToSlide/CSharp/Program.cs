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

namespace AddingPlainLineToSlide
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PresentationEx class that represents the PPTX file
            PresentationEx pres = new PresentationEx();

            //Get the first slide
            SlideEx sld = pres.Slides[0];

            //Add an autoshape of type line
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Line, 50, 150, 300, 0);

            //Write the PPTX to Disk
            pres.Write(dataDir + "LineShape1.pptx");

        }
    }
}