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
    public class AccessBuiltinProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instantiate the Presentation class that represents the presentation
            Presentation pres = new Presentation(dataDir + "AccessBuiltin Properties.pptx");

            //Create a reference to IDocumentProperties object associated with Presentation
            IDocumentProperties documentProperties = pres.DocumentProperties;

            //Display the builtin properties
            System.Console.WriteLine("Category : " + documentProperties.Category);
            System.Console.WriteLine("Current Status : " + documentProperties.ContentStatus);
            System.Console.WriteLine("Creation Date : " + documentProperties.CreatedTime);
            System.Console.WriteLine("Author : " + documentProperties.Author);
            System.Console.WriteLine("Description : " + documentProperties.Comments);
            System.Console.WriteLine("KeyWords : " + documentProperties.Keywords);
            System.Console.WriteLine("Last Modified By : " + documentProperties.LastSavedBy);
            System.Console.WriteLine("Supervisor : " + documentProperties.Manager);
            System.Console.WriteLine("Modified Date : " + documentProperties.LastSavedTime);
            System.Console.WriteLine("Presentation Format : " + documentProperties.PresentationFormat);
            System.Console.WriteLine("Last Print Date : " + documentProperties.LastPrinted);
            System.Console.WriteLine("Is Shared between producers : " + documentProperties.SharedDoc);
            System.Console.WriteLine("Subject : " + documentProperties.Subject);
            System.Console.WriteLine("Title : " + documentProperties.Title);
            
        }
    }
}