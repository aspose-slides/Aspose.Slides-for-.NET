//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace FromOnePositiontoAnotherwithinSamePresentation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "Aspose.pptx"))
            {

                //Clone the desired slide to the end of the collection of slides in the same presentation
                ISlideCollection slds = pres.Slides;

                //Clone the desired slide to the specified index in the same presentation
                slds.InsertClone(2, pres.Slides[1]);

                //Write the modified presentation to disk
                pres.Save(dataDir + "Aspose_clone.pptx", SaveFormat.Pptx);

            }
            
        }
    }
}