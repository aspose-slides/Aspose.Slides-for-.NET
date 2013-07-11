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

namespace AddingArrowShapedLineToSlide
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
            PresentationEx pres = new PresentationEx();

            //Get the first slide
            SlideEx sld = pres.Slides[0];

            //Add an autoshape of type line
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Line, 50, 150, 300, 0);
            ShapeEx shp = sld.Shapes[idx];

            //Apply some formatting on the line
            shp.LineFormat.Style = LineStyleEx.ThickBetweenThin;
            shp.LineFormat.Width = 10;

            shp.LineFormat.DashStyle = LineDashStyleEx.DashDot;

            shp.LineFormat.BeginArrowheadLength = LineArrowheadLengthEx.Short;
            shp.LineFormat.BeginArrowheadStyle = LineArrowheadStyleEx.Oval;

            shp.LineFormat.EndArrowheadLength = LineArrowheadLengthEx.Long;
            shp.LineFormat.EndArrowheadStyle = LineArrowheadStyleEx.Triangle;

            shp.LineFormat.FillFormat.FillType = FillTypeEx.Solid;
            shp.LineFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Maroon;

            //Write the PPTX to Disk
            pres.Write(dataDir + "LineShape.pptx");

        }
    }
}