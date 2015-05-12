//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ReplacingText
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate Presentation class that represents PPTX//Instantiate Presentation class that represents PPTX
            using (Presentation pres = new Presentation(dataDir + "Aspose.pptx"))
            {

                //Access first slide
                ISlide sld = pres.Slides[0];

                //Iterate through shapes to find the placeholder
                foreach (IShape shp in sld.Shapes)
                    if (shp.Placeholder != null)
                    {
                        //Change the text of each placeholder
                        ((IAutoShape)shp).TextFrame.Text = "This is Placeholder";
                    }

                //Save the PPTX to Disk
                pres.Save(dataDir + "output.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }

        }
    }
}