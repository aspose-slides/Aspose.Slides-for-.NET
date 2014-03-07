//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace RemovingWriteProtection
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Opening the presentation file
            Presentation pres = new Presentation(dataDir + "demoWriteProtected.pptx");


            //Checking if presentation is write protected
            if (pres.ProtectionManager.IsWriteProtected)
                //Removing Write protection
                pres.ProtectionManager.RemoveWriteProtection();

            //Saving presentation
            pres.Write(dataDir + "newDemo.pptx");
        }
    }
}