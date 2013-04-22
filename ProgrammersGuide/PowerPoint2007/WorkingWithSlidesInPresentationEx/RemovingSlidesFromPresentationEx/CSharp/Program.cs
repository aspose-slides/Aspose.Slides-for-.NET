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

namespace RemovingSlidesFromPresentationEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a PresentationEx object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            //Accessing a slide using its index in the slides collection
            SlideEx slide = pres.Slides[0];

            //Removing a slide using its reference
            pres.Slides.Remove(slide);

            //Writing the presentation as a PPTX file
            pres.Write(dataDir + "modified.pptx");
            

        }
    }
}