//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AccessingModifyingProperties
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instanciate the Presentation class that represents the PPTX
            Presentation pres = new Presentation(dataDir + "Aspose.pptx");

            //Create a reference to DocumentProperties object associated with Prsentation
            IDocumentProperties dp = pres.DocumentProperties;


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
            pres.Write(dataDir + "CustomDemoModified.pptx");
        }
    }
}