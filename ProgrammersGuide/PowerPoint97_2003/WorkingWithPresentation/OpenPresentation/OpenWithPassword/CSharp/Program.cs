//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace OpenWithPassword
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Creating instance of load options to set the presentation access password
            Aspose.Slides.LoadOptions loadOptions = new Aspose.Slides.LoadOptions();

            //Setting the access password
            loadOptions.Password = "123456";

            //Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "simplePasswordProtected.ppt", loadOptions);

            //Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString());
        }
    }
}