/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
*/

using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.Presentations
{
    public class Convert_XPS_Options
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "Convert_XPS_Options.pptx"))
            {
                // Instantiate the TiffOptions class
                XpsOptions opts = new XpsOptions();

                // Save MetaFiles as PNG
                opts.SaveMetafilesAsPng = true;

                // Save the presentation to XPS document
                pres.Save(dataDir + "XPS_With_Options.xps", SaveFormat.Xps, opts);
            }
        }
    }
}