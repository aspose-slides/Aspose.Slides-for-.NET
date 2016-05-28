//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.Presentations
{
    public class Convert_HTML
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "Convert_HTML.pptx"))
            {
                HtmlOptions htmlOpt = new HtmlOptions();
                htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

                //Saving the presentation to HTML
                presentation.Save(dataDir + "demo.html", Aspose.Slides.Export.SaveFormat.Html, htmlOpt);
            }
        }
    }
}