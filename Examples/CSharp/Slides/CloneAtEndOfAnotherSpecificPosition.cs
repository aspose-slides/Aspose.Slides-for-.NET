//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.Slides
{
    public class CloneAtEndOfAnotherSpecificPosition
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate Presentation class to load the source presentation file
            using (Presentation srcPres = new Presentation(dataDir + "CloneAtEndOfAnotherSpecificPosition.pptx"))
            {
                //Instantiate Presentation class for destination presentation (where slide is to be cloned)
                using (Presentation destPres = new Presentation())
                {
                    //Clone the desired slide from the source presentation to the end of the collection of slides in destination presentation
                    ISlideCollection slds = destPres.Slides;

                    //Clone the desired slide from the source presentation to the specified position in destination presentation
                    slds.InsertClone(1, srcPres.Slides[1]);


                    //Write the destination presentation to disk
                    destPres.Save(dataDir + "Aspose_out.pptx", SaveFormat.Pptx);
                }
            }
        }
    }
}