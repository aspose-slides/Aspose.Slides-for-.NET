using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example demonstrates how the custom RootDirectoryClsid can be set.
Note. Object class GUID (CLSID) that is stored in the root directory entry can be used for COM activation of the document's application.
The default value is '64818D11-4F9B-11CF-86EA-00AA00B929E8' that corresponds to 'Microsoft Powerpoint.Slide.8'.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class RootDirectoryClsId
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "pres.ppt");

            using (Presentation pres = new Presentation())
            {
                PptOptions pptOptions = new PptOptions();

                // set CLSID to 'Microsoft Powerpoint.Show.8'
                pptOptions.RootDirectoryClsid = new Guid("64818D10-4F9B-11CF-86EA-00AA00B929E8");

                // Save presentation
                pres.Save(resultPath, SaveFormat.Ppt, pptOptions);
            }
        }
    }
}
