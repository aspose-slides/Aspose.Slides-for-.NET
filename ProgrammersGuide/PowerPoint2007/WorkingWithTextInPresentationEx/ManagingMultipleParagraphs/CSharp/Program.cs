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

namespace ManagingMultipleParagraphs
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

            //Instantiate a PresentationEx class that represents a PPTX file
            PresentationEx pres = new PresentationEx();


            //Accessing first slide
            SlideEx slide = pres.Slides[0];

            //Add an AutoShape of Rectangle type
            int idx = slide.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 50, 150, 300, 150);
            AutoShapeEx ashp = (AutoShapeEx)slide.Shapes[idx];

            //Access TextFrame of the AutoShape
            TextFrameEx tf = ashp.TextFrame;

            //Create Paragraphs and Portions with different text formats
            ParagraphEx para0 = tf.Paragraphs[0];
            PortionEx port01 = new PortionEx();
            PortionEx port02 = new PortionEx();
            para0.Portions.Add(port01);
            para0.Portions.Add(port02);

            ParagraphEx para1 = new ParagraphEx();
            tf.Paragraphs.Add(para1);
            PortionEx port10 = new PortionEx();
            PortionEx port11 = new PortionEx();
            PortionEx port12 = new PortionEx();
            para1.Portions.Add(port10);
            para1.Portions.Add(port11);
            para1.Portions.Add(port12);

            ParagraphEx para2 = new ParagraphEx();
            tf.Paragraphs.Add(para2);
            PortionEx port20 = new PortionEx();
            PortionEx port21 = new PortionEx();
            PortionEx port22 = new PortionEx();
            para2.Portions.Add(port20);
            para2.Portions.Add(port21);
            para2.Portions.Add(port22);

            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 3; j++)
                {
                    tf.Paragraphs[i].Portions[j].Text = "Portion0" + j.ToString();
                    if (j == 0)
                    {
                        tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
                        tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.SolidFillColor.Color = Color.Red;
                        tf.Paragraphs[i].Portions[j].PortionFormat.FontBold = NullableBool.True;
                        tf.Paragraphs[i].Portions[j].PortionFormat.FontHeight = 15;
                    }
                    else if (j == 1)
                    {
                        tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
                        tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;
                        tf.Paragraphs[i].Portions[j].PortionFormat.FontItalic = NullableBool.True;
                        tf.Paragraphs[i].Portions[j].PortionFormat.FontHeight = 18;
                    }
                }

            //Write PPTX to Disk
            pres.Write(dataDir + "output.pptx");


        }
    }
}