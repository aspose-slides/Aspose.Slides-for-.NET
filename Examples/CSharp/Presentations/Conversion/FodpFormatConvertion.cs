using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The example demonstrates loading and saving presentation in Fodp format.
*/
namespace CSharp.Presentations.Conversion
{
    class FodpFormatConvertion
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Conversion();
            string outFodpPath = Path.Combine(RunExamples.OutPath, "FodpFormatConvertion.fodp");
            string outPptxPath = Path.Combine(RunExamples.OutPath, "FodpFormatConvertion.pptx");

            using (Presentation presentation = new Presentation(dataDir + "Example.fodp"))
            {
                presentation.Save(outPptxPath, SaveFormat.Pptx);
            }

            using (Presentation pres = new Presentation(outPptxPath))
            {
                pres.Save(outFodpPath, SaveFormat.Fodp);
            }
        }
    }
}
