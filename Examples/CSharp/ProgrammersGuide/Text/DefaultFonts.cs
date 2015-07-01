//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing.Imaging;
using Aspose.Slides.Export;

namespace CSharp.Text
{
    public class DefaultFonts
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            //Use load options to define the default regualr and asian fonts//Use load options to define the default regualr and asian fonts
            LoadOptions lo = new LoadOptions(LoadFormat.Auto);
            lo.DefaultRegularFont = "Wingdings";
            lo.DefaultAsianFont = "Wingdings";

            //Load the presentation
            using (Presentation pptx = new Presentation(dataDir+ "DefaultFonts.pptx", lo))
            {

                //Generate slide thumbnail
                pptx.Slides[0].GetThumbnail(1, 1).Save(dataDir+ "output.png", ImageFormat.Png);

                //Generate PDF
                pptx.Save("output.pdf", SaveFormat.Pdf);

                //Generate XPS
                pptx.Save(dataDir+ "output.xps", SaveFormat.Xps);
            }

            
            
        }
    }
}