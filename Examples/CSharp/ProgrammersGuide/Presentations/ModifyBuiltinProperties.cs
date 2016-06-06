//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

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