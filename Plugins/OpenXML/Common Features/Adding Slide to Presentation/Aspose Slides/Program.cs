// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate PresentationEx class that represents the PPT file
            Presentation pres = new Presentation();

            //Blank slide is added by default, when you create
            //presentation from default constructor
            //Adding an empty slide to the presentation and getting the reference of
            //that empty slide
            Slide slide = pres.AddEmptySlide();

            slide = pres.GetSlideByPosition(2);

            //Get the font index for Arial 
            //It is always 0 if you create presentation from
            //default constructor
            int arialFontIndex = 0;

            //Add a textbox
            //To add it, we will first add a rectangle
            Shape shp = slide.Shapes.AddRectangle(1200, 800, 3200, 370);

            //Hide its line
            shp.LineFormat.ShowLines = false;

            //Then add a textframe inside it
            TextFrame tf = shp.AddTextFrame("");

            //Set a text
            tf.Text = "Text added dynamically";
            Portion port = tf.Paragraphs[0].Portions[0];
            port.FontIndex = arialFontIndex;
            port.FontBold = true;
            port.FontHeight = 32;

            //Write the output to disk
            pres.Write("outAspose.ppt");

        }
    }
}