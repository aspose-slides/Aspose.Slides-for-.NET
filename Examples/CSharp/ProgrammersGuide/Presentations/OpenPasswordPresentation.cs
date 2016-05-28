//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace CSharp.Presentations
{
    public class OpenPasswordPresentation
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // creating instance of load options to set the presentation access password
            LoadOptions loadOptions = new LoadOptions();

            // Setting the access password
            loadOptions.Password = "pass";

            // Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "OpenPasswordPresentation.pptx", loadOptions);

            // Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString());
        }
    }
}