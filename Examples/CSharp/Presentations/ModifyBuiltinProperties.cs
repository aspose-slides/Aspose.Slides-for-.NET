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
    public class ModifyBuiltinProperties
    {
        public static void Run()
        {
            // For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate the Presentation class that represents the Presentation
            Presentation presentation = new Presentation(dataDir + "ModifyBuiltinProperties.pptx");

            // Create a reference to IDocumentProperties object associated with Presentation
            IDocumentProperties documentProperties = presentation.DocumentProperties;

            //Set the builtin properties
            documentProperties.Author = "Aspose.Slides for .NET";
            documentProperties.Title = "Modifying Presentation Properties";
            documentProperties.Subject = "Aspose Subject";
            documentProperties.Comments = "Aspose Description";
            documentProperties.Manager = "Aspose Manager";

            //Save your presentation to a file
            presentation.Save(dataDir + "DocumentProperties.pptx", SaveFormat.Pptx);
            
        }
    }
}