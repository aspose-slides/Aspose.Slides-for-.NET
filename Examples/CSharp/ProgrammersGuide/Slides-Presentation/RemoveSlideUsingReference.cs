//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace CSharp.Slides
{
    public class RemoveSlideUsingReference
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "RemoveSlideUsingReference.pptx"))
            {

                //Accessing a slide using its index in the slides collection
                ISlide slide = pres.Slides[0];


                //Removing a slide using its reference
                pres.Slides.Remove(slide);


                //Writing the presentation file
                pres.Save(dataDir + "modified.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}