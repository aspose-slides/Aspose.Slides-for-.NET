//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
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
    public class PPTtoPPTX
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instantiate a Presentation object that represents a PPTX file
            Presentation pres = new Presentation(dataDir + "PPTtoPPTX.ppt");

            //Saving the PPTX presentation to PPTX format
            pres.Save(dataDir + "PPTtoPPTX.pptx", SaveFormat.Pptx);


        }
    }
}