//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ImportingHTMLText
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Create Empty presentation instance//Create Empty presentation instance
            using (Presentation pres = new Presentation())
            {
                //Acesss the default first slide of presentation
                ISlide slide = pres.Slides[0];

                //Adding the AutoShape to accomodate the HTML content
                IAutoShape ashape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, pres.SlideSize.Size.Width - 20, pres.SlideSize.Size.Height - 10);

                ashape.FillFormat.FillType = FillType.NoFill;

                //Adding text frame to the shape
                ashape.AddTextFrame("");

                //Clearing all paragraphs in added text frame
                ashape.TextFrame.Paragraphs.Clear();

                //Loading the HTML file using stream reader
                TextReader tr = new StreamReader(dataDir + "file.html");

                //Adding text from HTML stream reader in text frame
                ashape.TextFrame.Paragraphs.AddFromHtml(tr.ReadToEnd());

                //Saving Presentation
                pres.Save(dataDir + "output.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
 
        }
    }
}