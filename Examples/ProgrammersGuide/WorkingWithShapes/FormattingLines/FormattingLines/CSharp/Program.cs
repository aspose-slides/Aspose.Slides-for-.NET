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
using Aspose.Slides.Export;

namespace FormattingLines
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

            //Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                //Get the first slide
                ISlide sld = pres.Slides[0];

                //Add autoshape of rectangle type
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 75);

                //Set the fill color of the rectangle shape
                shp.FillFormat.FillType = FillType.Solid;
                shp.FillFormat.SolidFillColor.Color = Color.White;

                //Apply some formatting on the line of the rectangle
                shp.LineFormat.Style = LineStyle.ThickThin;
                shp.LineFormat.Width = 7;
                shp.LineFormat.DashStyle = LineDashStyle.Dash;

                //set the color of the line of rectangle
                shp.LineFormat.FillFormat.FillType = FillType.Solid;
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;

                //Write the PPTX file to disk
                pres.Save(dataDir + "RectShpLn.pptx", SaveFormat.Pptx);

            } 
        }
    }
}