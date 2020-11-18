using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using DataTable = System.Data.DataTable;

namespace CSharp.Presentations.Conversion
{
    // This example demonstrates setting keep text out of 3D scene.

    public class KeepTextFlat
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Text();
            string pptxFileName = Path.Combine(dataDir, "KeepTextFlat.pptx");
            string resultPath = Path.Combine(RunExamples.OutPath, "KeepTextFlat_out.png");

            using (Presentation pres = new Presentation(pptxFileName))
            {
                var shape1 = pres.Slides[0].Shapes[0] as AutoShape;
                var shape2 = pres.Slides[0].Shapes[1] as AutoShape;

                shape1.TextFrame.TextFrameFormat.KeepTextFlat = false;
                shape2.TextFrame.TextFrameFormat.KeepTextFlat = true;

                pres.Slides[0].GetThumbnail(4 / 3f, 4 / 3f).Save(resultPath, ImageFormat.Png);
            }
        }
    }
}
