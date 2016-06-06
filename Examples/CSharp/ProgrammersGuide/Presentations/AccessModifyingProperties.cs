//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

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
    public class AccessModifyingProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instanciate the Presentation class that represents the PPTX
            Presentation presentation = new Presentation(dataDir + "AccessModifyingProperties.pptx");

            //Create a reference to DocumentProperties object associated with Prsentation
            IDocumentProperties documentProperties = presentation.DocumentProperties;

            //Access and modify custom properties
            for (int i = 0; i < documentProperties.CountOfCustomProperties; i++)
            {
                //Display names and values of custom properties
                System.Console.WriteLine("Custom Property Name : " + documentProperties.GetCustomPropertyName(i));
                System.Console.WriteLine("Custom Property Value : " + documentProperties[documentProperties.GetCustomPropertyName(i)]);

                //Modify values of custom properties
                documentProperties[documentProperties.GetCustomPropertyName(i)] = "New Value " + (i + 1);
            }

            //Save your presentation to a file
            presentation.Save(dataDir + "CustomDemoModified.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}