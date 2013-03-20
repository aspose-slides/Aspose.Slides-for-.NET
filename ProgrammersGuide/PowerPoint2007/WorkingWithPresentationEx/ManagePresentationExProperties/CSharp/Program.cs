//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;

namespace ManagePresentationExProperties
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate the PresentationEx class that represents the PPTX
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            //Create a reference to DocumentProperties object associated with PrsentationEx
            DocumentPropertiesEx dp = pres.DocumentProperties;

            // 1.
            //Display the built in properties
            System.Console.WriteLine("Category : " + dp.Category);
            System.Console.WriteLine("Current Status : " + dp.ContentStatus);
            System.Console.WriteLine("Creation Date : " + dp.Created);
            System.Console.WriteLine("Author : " + dp.Creator);
            System.Console.WriteLine("Description : " + dp.Description);
            System.Console.WriteLine("KeyWords : " + dp.Keywords);
            System.Console.WriteLine("Last Modified By : " + dp.LastModifiedBy);
            System.Console.WriteLine("Supervisor : " + dp.Manager);
            System.Console.WriteLine("Modified Date : " + dp.Modified);
            System.Console.WriteLine("Presentation Format : " + dp.PresentationFormat);
            System.Console.WriteLine("Last Print Date : " + dp.Printed);
            System.Console.WriteLine("Is Shared between producers : " + dp.SharedDoc);
            System.Console.WriteLine("Subject : " + dp.Subject);
            System.Console.WriteLine("Title : " + dp.Title);


            // 2.
            //Set the built in properties
            dp.Creator = "Aspose.Slides for .NET";
            dp.Title = "Modifying Presentation Properties";
            dp.Subject = "Aspose Subject";
            dp.Description = "Aspose Description";
            dp.Manager = "Aspose Manager";

            //Save your presentation to a file
            pres.Write(dataDir + "updatedProperties.pptx");


            // 3.
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
            pres.Write(dataDir + "updatedCustomProperties.pptx");
        }
    }
}