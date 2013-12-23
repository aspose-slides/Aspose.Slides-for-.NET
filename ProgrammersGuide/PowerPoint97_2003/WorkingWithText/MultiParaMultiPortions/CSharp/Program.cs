//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing;

namespace MultiParaMultiPortions
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

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation();
            
            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(1);

            //Adding Rectangle Shape and setting its properties
            Aspose.Slides.Rectangle rect = slide.Shapes.AddRectangle(50, 50, 500, 200);
            LineFormat lf = rect.LineFormat;
            lf.ShowLines = false;
            TextFrame tf = rect.AddTextFrame("Default Text");
            tf.FitShapeToText = true;

            //Creating Paragraphs and Portions
            Paragraph para0 = tf.Paragraphs[0];
            Portion port01 = new Portion();
            Portion port02 = new Portion();
            para0.Portions.Add(port01);
            para0.Portions.Add(port02);

            Paragraph para1 = new Paragraph();
            tf.Paragraphs.Add(para1);
            Portion port10 = new Portion();
            Portion port11 = new Portion();
            Portion port12 = new Portion();
            para1.Portions.Add(port10);
            para1.Portions.Add(port11);
            para1.Portions.Add(port12);

            Paragraph para2 = new Paragraph();
            tf.Paragraphs.Add(para2);
            Portion port20 = new Portion();
            Portion port21 = new Portion();
            Portion port22 = new Portion();
            para2.Portions.Add(port20);
            para2.Portions.Add(port21);
            para2.Portions.Add(port22);

            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 3; j++)
                {
                    tf.Paragraphs[i].Portions[j].Text = "Portion0" + j.ToString();
                    if (j == 0)
                        tf.Paragraphs[i].Portions[j].FontColor = Color.Red;
                    else if (j == 1)
                        tf.Paragraphs[i].Portions[j].FontColor = Color.Blue;
                }
            pres.Write(dataDir + "modified.ppt");
        }
    }
}