//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace OpenSimple
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Opening the presentation file by passing the file path to the constructor
            //of the Presentation class
            Presentation pres = new Presentation(dataDir + "simple.ppt");

            //Printing the total number of slides in the presentation
            System.Console.WriteLine("Number of slides in simple presentation are : " + pres.Slides.Count.ToString());
        }
    }
}