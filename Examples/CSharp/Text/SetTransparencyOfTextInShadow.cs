using Aspose.Slides;
using Aspose.Slides.Effects;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class SetTransparencyOfTextInShadow
    {
        public static void Run() {
            //ExStart:SetTransparencyOfTextInShadow
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            using (Presentation pres = new Presentation(dataDir+ "transparency.pptx"))
            {
                IAutoShape shape = (IAutoShape)pres.Slides[0].Shapes[0];
                IEffectFormat effects = shape.TextFrame.Paragraphs[0].Portions[0].PortionFormat.EffectFormat;

                IOuterShadow outerShadowEffect = effects.OuterShadowEffect;

                Color shadowColor = outerShadowEffect.ShadowColor.Color;
                Console.WriteLine("{0} - transparency is: {1}", shadowColor, ((float)shadowColor.A / byte.MaxValue) * 100);

                // set transparency to zero percent
                outerShadowEffect.ShadowColor.Color = Color.FromArgb(255, shadowColor);

                pres.Save(dataDir+"transparency-2.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SetTransparencyOfTextInShadow
        }
    }
}
