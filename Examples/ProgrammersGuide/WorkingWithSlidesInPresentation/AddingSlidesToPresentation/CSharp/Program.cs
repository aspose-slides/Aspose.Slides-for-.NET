//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AddingSlidesToPresentation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");


            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);


            //Instantiate Presentation class that represents the presentation file
            using (Presentation pres = new Presentation())
            {

                //Instantiate SlideCollection calss
                ISlideCollection slds = pres.Slides;

                for (int i = 0; i < pres.LayoutSlides.Count; i++)
                {
                    //Add an empty slide to the Slides collection
                    slds.AddEmptySlide(pres.LayoutSlides[i]);

                }
                //Do some work on the newly added slide

                //Save the PPTX file to the Disk
                pres.Write(dataDir + "EmptySlide.pptx");

            }
        }
    }
}