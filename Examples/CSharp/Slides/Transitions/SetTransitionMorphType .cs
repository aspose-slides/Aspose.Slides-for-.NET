using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Transitions
{
    class SetTransitionMorphType
    {
        public static void Run() {

            //ExStart:SetTransitionMorphType
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Transitions();

            using (Presentation presentation = new Presentation(dataDir + "presentation.pptx"))
            {
                presentation.Slides[0].SlideShowTransition.Type = TransitionType.Morph;
                ((IMorphTransition)presentation.Slides[0].SlideShowTransition.Value).MorphType = TransitionMorphType.ByWord;
                presentation.Save(dataDir + "presentation-out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SetTransitionMorphType
        }
    }
}
