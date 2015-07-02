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
    public class SaveProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation();

            //....do some work here.....

            //Setting access to document properties in password protected mode
            pres.ProtectionManager.EncryptDocumentProperties = false;

            //Setting Password
            pres.ProtectionManager.Encrypt("pass");

            //Save your presentation to a file
            pres.Save(dataDir + "demoPassDocument.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}