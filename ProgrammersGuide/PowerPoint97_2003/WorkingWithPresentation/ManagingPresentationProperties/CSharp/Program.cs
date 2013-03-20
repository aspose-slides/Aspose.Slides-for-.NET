//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ManagingPresentationProperties
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");

            //Create a reference to DocumentProperties associated with Presentation
            DocumentProperties dp = pres.DocumentProperties;

            //Accessing the built-in properties of the presentation
            System.Console.WriteLine("Author: " + dp.Author);
            System.Console.WriteLine("Title: " + dp.Title);
            System.Console.WriteLine("Company: " + dp.Company);
            System.Console.WriteLine("Comments: " + dp.Comments);
            System.Console.WriteLine("Subject: " + dp.Subject);

            System.Console.WriteLine("");
            System.Console.WriteLine("Updating presentation properties now ");
            System.Console.WriteLine("");

            //Modifying the built-in properties of the presentation
            dp.Author = "Mudassir Fayyaz";
            dp.Title = "Modifying Presentation Properties";
            dp.Company = "Aspose Pty. Ltd.";
            dp.Comments = "Modified Presentation Properties";
            dp.Subject = "Presentation Properties";


            //Save your presentation to a file
            pres.Write(dataDir + "modified.ppt");

            //Access and modify custom properties
            for (int i = 0; i < dp.Count; i++)
            {
                //Display names and values of custom properties
                System.Console.WriteLine("Custom Property Name : " + dp.GetPropertyName(i));
                System.Console.WriteLine("Custom Property Value : " + dp[dp.GetPropertyName(i)]);

                //Modify values of custom properties
                dp[dp.GetPropertyName(i)] = "New Value " + (i + 1);
            }

            //Save your presentation to a file
            pres.Write(dataDir + "DemoProps.ppt");
        }
    }
}