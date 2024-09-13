using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;

/*
This example demonstrates hao to get raw text of a presentation using PresentationFactory.Instance.GetPresentationText method.
*/

namespace CSharp.Text
{
    class GetPresentationRowTextExample
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "PresentationText.pptx");

            IPresentationText presentationText = PresentationFactory.Instance.GetPresentationText(presentationName, TextExtractionArrangingMode.Unarranged);

            Console.WriteLine("Comments for Slide 1 : {0}", presentationText.SlidesText[0].CommentsText);
            Console.WriteLine("Text for Slide 1 : {0}", presentationText.SlidesText[0].Text);
            Console.WriteLine("Notes for Slide 2 : {0}", presentationText.SlidesText[1].NotesText);
        }
    }
}
