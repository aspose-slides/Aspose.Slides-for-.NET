//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace SettingBackgroundColorToGradientToSlides
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate the Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "Aspose.pptx"))
            {

                //Apply Gradiant effect to the Background
                pres.Slides[0].Background.Type = BackgroundType.OwnBackground;
                pres.Slides[0].Background.FillFormat.FillType = FillType.Gradient;
                pres.Slides[0].Background.FillFormat.GradientFormat.TileFlip = TileFlip.FlipBoth;

                //Write the presentation to disk
                pres.Save(dataDir + "ContentBG_Grad.pptx", SaveFormat.Pptx)
                ;
            }
 
        }
    }
}