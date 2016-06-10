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
    public class AccessProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Accessing the Document Properties of a Password Protected Presentation without Password
            // creating instance of load options to set the presentation access password
            LoadOptions loadOptions = new LoadOptions();

            // Setting the access password to null
            loadOptions.Password = null;

            // Setting the access to document properties
            loadOptions.OnlyLoadDocumentProperties = true;

            // Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "AccessProperties.pptx", loadOptions);

            //Getting Document Properties
            IDocumentProperties docProps = pres.DocumentProperties;

            System.Console.WriteLine("Name of Application : " + docProps.NameOfApplication);
        }
    }
}