//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;
using System.Drawing;

namespace ManagingFontFamilyOfText
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate PresentationEx Class
            PresentationEx pres = new PresentationEx();


            //Get first slide
            SlideEx sld = pres.Slides[0];

            //Add an AutoShape of Rectangle type
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 50, 50, 200, 50);

            //Access the added AutoShape
            AutoShapeEx ashp = (AutoShapeEx)sld.Shapes[idx];

            //Remove any fill style associated with the AutoShape
            ashp.FillFormat.FillType = FillTypeEx.NoFill;

            //Access the TextFrame associated with the AutoShape
            TextFrameEx tf = ashp.TextFrame;
            tf.Text = "Aspose TextBox";

            //Access the Portion associated with the TextFrame
            PortionEx port = tf.Paragraphs[0].Portions[0];

            //Set the Font for the Portion
            port.PortionFormat.LatinFont = new FontDataEx("Times New Roman");

            //Set Bold property of the Font 
            port.PortionFormat.FontBold = NullableBool.True;

            //Set Italic property of the Font
            port.PortionFormat.FontItalic = NullableBool.True;

            //Set Underline property of the Font 
            port.PortionFormat.FontUnderline = TextUnderlineTypeEx.Single;

            //Set the Height of the Font
            port.PortionFormat.FontHeight = 25;

            //Set the color of the Font
            port.PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
            port.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;

            //Write the presentation to disk
            pres.Write(dataDir + "output.pptx");


        }
    }
}