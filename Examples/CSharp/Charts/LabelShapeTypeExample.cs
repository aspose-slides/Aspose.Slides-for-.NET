using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to set the shape in a chart series.
 */

namespace CSharp.Charts
{
    public class LabelShapeTypeExample
    {
        public static void Run()
        {
            // Source file name
            string sourcePath = Path.Combine(RunExamples.GetDataDir_Charts(), "LabelShapeType.pptx");
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "LabelShapeType-out.pptx");

            using (Presentation pres = new Presentation(sourcePath))
            {
                var lineSerie = (pres.Slides[0].Shapes[1] as IChart).ChartData.Series[0];

                var label0 = lineSerie.Labels[0];
                label0.LabelShapeType = ShapeType.Ellipse;
                label0.DataLabelFormat.Format.Fill.FillType = FillType.Solid;
                label0.DataLabelFormat.Format.Fill.SolidFillColor.Color = Color.Blue;

                var label1 = lineSerie.Labels[1];
                label1.LabelShapeType = ShapeType.Cloud;
                label1.DataLabelFormat.Format.Fill.FillType = FillType.Solid;
                label1.DataLabelFormat.Format.Fill.SolidFillColor.Color = Color.Green;

                var label2 = lineSerie.Labels[2];
                label2.LabelShapeType = ShapeType.EqualMath;
                label2.DataLabelFormat.Format.Fill.FillType = FillType.Solid;
                label2.DataLabelFormat.Format.Fill.SolidFillColor.Color = Color.Yellow;

                var label3 = lineSerie.Labels[3];
                label3.LabelShapeType = ShapeType.CalloutDownArrow;
                label3.DataLabelFormat.Format.Fill.FillType = FillType.Solid;
                label3.DataLabelFormat.Format.Fill.SolidFillColor.Color = Color.Red;

                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}