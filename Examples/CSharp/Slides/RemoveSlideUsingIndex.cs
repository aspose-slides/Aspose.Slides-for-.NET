//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace CSharp.Slides
{
    public class RemoveSlideUsingIndex
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "RemoveSlideUsingIndex.pptx"))
            {

                //Removing a slide using its slide index
                pres.Slides.RemoveAt(0);


                //Writing the presentation file
                pres.Save(dataDir + "modified.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
        }
    }
}