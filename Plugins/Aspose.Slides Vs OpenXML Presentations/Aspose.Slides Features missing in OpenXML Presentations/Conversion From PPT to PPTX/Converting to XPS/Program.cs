// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;
namespace Converting_to_XPS
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.ppt");
            //Saving the presentation to TIFF document
            pres.Save(MyDir + "converted.xps", Aspose.Slides.Export.SaveFormat.Xps);
        }
    }
}
