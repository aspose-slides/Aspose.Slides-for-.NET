//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.SlideShow;
using Aspose.Slides.Export;

namespace CSharp.Slides
{
    public class SimpleSlideTransitions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "SimpleSlideTransitions.pptx"))
            {

                //Apply circle type transition on slide 1
                pres.Slides[0].SlideShowTransition.Type = TransitionType.Circle;

                //Apply comb type transition on slide 2
                pres.Slides[1].SlideShowTransition.Type = TransitionType.Comb;

                //Write the presentation to disk
                pres.Save(dataDir + "SampleTransition.pptx", SaveFormat.Pptx);

            }
        }
    }
}