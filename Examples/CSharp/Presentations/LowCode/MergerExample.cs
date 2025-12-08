using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;
using Aspose.Slides.LowCode;

/*
This example shows how to merge the set of input presentations of the same format into a single presentation file.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.LowCode
{
    class MergerExample
    {
        public static void Run()
        {
            string pptxFileName1 = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_LowCode(), "ForEachPortion.pptx");
            string pptxFileName2 = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_LowCode(), "ConvertExample.pptx");
            string pptxFileName3 = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_LowCode(), "MultipleMaster.pptx");
            string outPpptxFile = Path.Combine(RunExamples.OutPath, "Merged-out.pptx");

            Merger.Process(new string[] { pptxFileName1, pptxFileName2, pptxFileName3 }, outPpptxFile);
        }
    }
}