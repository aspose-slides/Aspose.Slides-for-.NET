//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace CSharp.Presentations
{
    public class OpenPresentation
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Opening the presentation file by passing the file path to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "OpenPresentation.pptx");

            //Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString());
        }
    }
}