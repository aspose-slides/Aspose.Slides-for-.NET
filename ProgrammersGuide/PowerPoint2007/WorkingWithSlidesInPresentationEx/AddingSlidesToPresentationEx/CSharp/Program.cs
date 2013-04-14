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

namespace AddingSlidesToPresentationEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PresentationEx class that represents the PPTX file
            PresentationEx pres = new PresentationEx();

            //Instantiate SlideExCollection calss
            SlideExCollection slds = pres.Slides;

            //Add an empty slide to the SlidesEx collection
            slds.AddEmptySlide(pres.LayoutSlides[0]);

            //Do some work on the newly added slide

            //Save the PPTX file to the Disk
            pres.Write(dataDir + "EmptySlide.pptx");

            
        }
    }
}