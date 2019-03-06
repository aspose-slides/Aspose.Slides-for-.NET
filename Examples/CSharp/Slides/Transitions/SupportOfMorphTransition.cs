using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Transitions
{
    class SupportOfMorphTransition
    {
        public static void Run()
        {
            //ExStart:SupportOfMorphTransition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Transitions();

            using (Presentation presentation = new Presentation())
            {
                AutoShape autoshape = (AutoShape)presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 400, 100);
                autoshape.TextFrame.Text = "Test text";

                presentation.Slides.AddClone(presentation.Slides[0]);

                presentation.Slides[1].Shapes[0].X += 100;
                presentation.Slides[1].Shapes[0].Y += 50;
                presentation.Slides[1].Shapes[0].Width -= 200;
                presentation.Slides[1].Shapes[0].Height -= 10;

                presentation.Slides[1].SlideShowTransition.Type = Aspose.Slides.SlideShow.TransitionType.Morph;

                presentation.Save(dataDir+"presentation-out.pptx", SaveFormat.Pptx);
            }

            //ExEnd:SupportOfMorphTransition
        }
    }
}
