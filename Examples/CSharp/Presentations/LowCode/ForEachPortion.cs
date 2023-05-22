using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;
using Aspose.Slides.LowCode;

/*
This code demonstrates how to iterate through each portion in the Presentation 
and get the text of the portions contained only in the notes slide.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.LowCode
{
    class ForEachPortion
    {
        public static void Run()
        {
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_LowCode(), "ForEachPortion.pptx");

            using (Presentation pres = new Presentation(pptxFileName))
            {
                ForEach.Portion(pres, true, (portion, para, slide, index) =>
                {
                    if (slide is NotesSlide && !string.IsNullOrEmpty(portion.Text))
                        System.Console.WriteLine($"{portion.Text}, index: {index}");
                });
            }
        }
    }
}

