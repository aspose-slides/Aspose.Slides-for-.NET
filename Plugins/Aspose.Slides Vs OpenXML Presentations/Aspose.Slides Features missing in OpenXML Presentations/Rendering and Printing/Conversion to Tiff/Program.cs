// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Conversion_to_Tiff
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a presentation file
            using (PresentationEx pres = new PresentationEx(MyDir + "Conversion.pptx"))
            {

                //Saving the presentation to TIFF document
                pres.Save(MyDir + "Converted.tiff", Aspose.Slides.Export.SaveFormat.Tiff);
            }
        }
    }
}
