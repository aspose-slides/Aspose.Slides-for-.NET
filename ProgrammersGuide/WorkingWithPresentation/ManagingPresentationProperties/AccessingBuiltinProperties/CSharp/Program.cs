//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AccessingBuiltinProperties
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate the Presentation class that represents the presentation
            Presentation pres = new Presentation(dataDir + "Aspose.pptx");

            //Create a reference to IDocumentProperties object associated with Presentation
            IDocumentProperties dp = pres.DocumentProperties;

            //Display the builtin properties
            System.Console.WriteLine("Category : " + dp.Category);
            System.Console.WriteLine("Current Status : " + dp.ContentStatus);
            System.Console.WriteLine("Creation Date : " + dp.CreatedTime);
            System.Console.WriteLine("Author : " + dp.Author);
            System.Console.WriteLine("Description : " + dp.Comments);
            System.Console.WriteLine("KeyWords : " + dp.Keywords);
            System.Console.WriteLine("Last Modified By : " + dp.LastSavedBy);
            System.Console.WriteLine("Supervisor : " + dp.Manager);
            System.Console.WriteLine("Modified Date : " + dp.LastSavedTime);
            System.Console.WriteLine("Presentation Format : " + dp.PresentationFormat);
            System.Console.WriteLine("Last Print Date : " + dp.LastPrinted);
            System.Console.WriteLine("Is Shared between producers : " + dp.SharedDoc);
            System.Console.WriteLine("Subject : " + dp.Subject);
            System.Console.WriteLine("Title : " + dp.Title);
            
        }
    }
}