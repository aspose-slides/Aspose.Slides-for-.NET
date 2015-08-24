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
    public class AccessOpenDoc
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Open the ODP file
            Presentation pres = new Presentation(dataDir + "AccessOpenDoc.odp");

            //Saving the ODP presentation to PPTX format
            pres.Save(dataDir + "AccessOpenDoc.pptx", SaveFormat.Pptx);

        }
    }
}