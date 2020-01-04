using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.CRUD
{
    class CloneSlideIntoSpecifiedSection
    {
        public static void Run() {

            //ExStart:CloneSlideIntoSpecifiedSection

            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            using (IPresentation presentation = new Presentation()) {

                presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 200, 50, 300, 100);
                presentation.Sections.AddSection("Section 1", presentation.Slides[0]);

                ISection section2 = presentation.Sections.AppendEmptySection("Section 2");

                presentation.Slides.AddClone(presentation.Slides[0], section2);


                presentation.Save(dataDir + "CloneSlideIntoSpecifiedSection.pptx",SaveFormat.Pptx);
            }
            //ExEnd:CloneSlideIntoSpecifiedSection

        }


    }
}
