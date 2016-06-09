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
    public class RemoveWriteProtection
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Opening the presentation file
            Presentation presentation = new Presentation(dataDir + "RemoveWriteProtection.pptx");

            //Checking if presentation is write protected
            if (presentation.ProtectionManager.IsWriteProtected)
                //Removing Write protection                
                presentation.ProtectionManager.RemoveWriteProtection();

            //Saving presentation
            presentation.Save(dataDir + "File_Without_WriteProtection.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}