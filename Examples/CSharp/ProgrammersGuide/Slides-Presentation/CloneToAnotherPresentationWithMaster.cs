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
    public class CloneToAnotherPresentationWithMaster
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate Presentation class to load the source presentation file

            using (Presentation srcPres = new Presentation(dataDir + "CloneToAnotherPresentationWithMaster.pptx"))
            {

                //Instantiate Presentation class for destination presentation (where slide is to be cloned)

                using (Presentation destPres = new Presentation())
                {

                    //Instantiate ISlide from the collection of slides in source presentation along with
                    //master slide
                    ISlide SourceSlide = srcPres.Slides[0];
                    IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

                    //Clone the desired master slide from the source presentation to the collection of masters in the
                    //destination presentation
                    IMasterSlideCollection masters = destPres.Masters;
                    IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

                    //Clone the desired master slide from the source presentation to the collection of masters in the
                    //destination presentation
                    IMasterSlide iSlide = masters.AddClone(SourceMaster);

                    //Clone the desired slide from the source presentation with the desired master to the end of the
                    //collection of slides in the destination presentation
                    ISlideCollection slds = destPres.Slides;
                    slds.AddClone(SourceSlide, iSlide);
                    //Clone the desired master slide from the source presentation to the collection of masters in the//destination presentation
                    //Save the destination presentation to disk
                    destPres.Save(dataDir + "Aspose_out.pptx", SaveFormat.Pptx);

                }
            }
        }
    }
}