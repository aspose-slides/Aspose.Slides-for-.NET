using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Tables
{
    class LockAspectRatio
    {
        public static void Run()
        {
            //ExStart:LockAspectRatio
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            using (Presentation pres = new Presentation(dataDir + "pres.pptx"))
            {
                ITable table = (ITable)pres.Slides[0].Shapes[0];
                Console.WriteLine($"Lock aspect ratio set: {table.ShapeLock.AspectRatioLocked}");

                table.ShapeLock.AspectRatioLocked = !table.ShapeLock.AspectRatioLocked; // invert

                Console.WriteLine($"Lock aspect ratio set: {table.ShapeLock.AspectRatioLocked}");

                pres.Save(dataDir + "pres-out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:LockAspectRatio

        }
    }
}
