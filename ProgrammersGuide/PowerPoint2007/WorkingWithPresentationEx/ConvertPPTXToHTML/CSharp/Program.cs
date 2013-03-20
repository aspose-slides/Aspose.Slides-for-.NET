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

namespace ConvertPPTXToHTML
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a PresentationEx object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            //Saving the presentation to HTML
            pres.Save(dataDir + "demo.html", Aspose.Slides.Export.SaveFormat.Html);
        }
    }
}