// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Converting_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.ppt");
            //Save the presentation to PDF with default options
            pres.Save(MyDir + "Converted.pdf", Aspose.Slides.Export.SaveFormat.Pdf);
        }
    }
}
