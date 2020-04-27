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
The example shows of using read-only recommendation for presentation (this feature was introduced in PowerPoint 2019).
When enabled it makes PowerPoint to open the presentation in read-only mode first.
*/
namespace CSharp.Presentations.Properties
{
    class ReadOnlyRecommended
    {
        public static void Run()
        {
            string outPptxPath = Path.Combine(RunExamples.OutPath, "ReadOnlyRecommended.pptx");

            using (Presentation pres = new Presentation())
            {
                pres.ProtectionManager.ReadOnlyRecommended = true;
                pres.Save(outPptxPath, SaveFormat.Pptx);
            }
        }
    }
}
