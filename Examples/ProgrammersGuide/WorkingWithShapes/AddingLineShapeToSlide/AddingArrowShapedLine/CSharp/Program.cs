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
using System.Drawing;

namespace AddingArrowShapedLine
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

            //Instantiate PresentationEx class that represents the PPTX file
            using (Presentation pres = new Presentation())
            {

                //Get the first slide
                ISlide sld = pres.Slides[0];

                //Add an autoshape of type line
                IAutoShape shp = sld.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0);

                //Apply some formatting on the line
                shp.LineFormat.Style = LineStyle.ThickBetweenThin;
                shp.LineFormat.Width = 10;

                shp.LineFormat.DashStyle = LineDashStyle.DashDot;

                shp.LineFormat.BeginArrowheadLength = LineArrowheadLength.Short;
                shp.LineFormat.BeginArrowheadStyle = LineArrowheadStyle.Oval;

                shp.LineFormat.EndArrowheadLength = LineArrowheadLength.Long;
                shp.LineFormat.EndArrowheadStyle = LineArrowheadStyle.Triangle;

                shp.LineFormat.FillFormat.FillType = FillType.Solid;
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Maroon;

                //Write the PPTX to Disk
                pres.Save(dataDir + "LineShape2.pptx", SaveFormat.Pptx);
            }

        }
    }
}