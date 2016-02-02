// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Export;
using Aspose.Slides.Pptx;

namespace Conversion_PPT_to_PPTX
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a PPTX file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.ppt");
            //Saving the PPTX presentation to PPTX format
            pres.Save(MyDir + "Converted.pptx", SaveFormat.Pptx);
        }
    }
}
