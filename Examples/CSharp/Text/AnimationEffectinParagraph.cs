using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class AnimationEffectinParagraph
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            //ExStart:AnimationEffectinParagraph
            using (Presentation presentation = new Presentation(dataDir + "Presentation1.pptx"))
            {
                // select paragraph to add effect
                IAutoShape autoShape = (IAutoShape)presentation.Slides[0].Shapes[0];
                IParagraph paragraph = autoShape.TextFrame.Paragraphs[0];

                // add Fly animation effect to selected paragraph
                IEffect effect = presentation.Slides[0].Timeline.MainSequence.AddEffect(paragraph, EffectType.Fly, EffectSubtype.Left, EffectTriggerType.OnClick);


                presentation.Save(dataDir + "AnimationEffectinParagraph.pptx", SaveFormat.Pptx);
            }



            //ExEnd:AnimationEffectinParagraph
        }
    }
}
