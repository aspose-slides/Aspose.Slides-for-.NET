using System;
using System.Collections.Generic;
using System.Text;

namespace Conversion_from_PPt_to_PPtx_format
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            PresentationEx presentation = new PresentationEx(MyDir + "Sample.ppt");
            presentation.Save(MyDir + "Converted.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

        }
    }
}
