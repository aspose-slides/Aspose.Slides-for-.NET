using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Vba;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.VBA
{
    class ExtractingVBAMacros
    {
        public static void Run()
        {
            //ExStart:ExtractingVBAMacros

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_VBA();

            using (Presentation pres = new Presentation(dataDir + "VBA.pptm"))
            {
                if (pres.VbaProject != null) // check if Presentation contains VBA Project
                {
                    foreach (IVbaModule module in pres.VbaProject.Modules)
                    {
                        Console.WriteLine(module.Name);
                        Console.WriteLine(module.SourceCode);
                    }
                }
            }

            //ExEnd:ExtractingVBAMacros

        }
    }
}
