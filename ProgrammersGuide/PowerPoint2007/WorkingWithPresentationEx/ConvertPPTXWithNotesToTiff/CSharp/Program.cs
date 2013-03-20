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

namespace ConvertPPTXWithNotesToTiff
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a PresentationEx object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "TestNotes.pptx");

            //Saving the presentation to TIFF notes
            pres.Save(dataDir + "TestNotes.tiff", Aspose.Slides.Export.SaveFormat.TiffNotes);
        }
    }
}