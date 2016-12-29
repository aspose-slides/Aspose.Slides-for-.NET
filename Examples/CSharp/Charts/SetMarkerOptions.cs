using System;
using System.Drawing;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class SetMarkerOptions
    {
        public static void Run()
        {
            //ExStart:SetMarkerOptions
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            ISlide slide = presentation.Slides[0];

            // Creating the default chart
            IChart chart = slide.Shapes.AddChart(ChartType.LineWithMarkers, 0, 0, 400, 400);

            // Getting the default chart data worksheet index
            int defaultWorksheetIndex = 0;

            // Getting the chart data worksheet
            IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

            // Delete demo series
            chart.ChartData.Series.Clear();

            // Add new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, "Series 1"), chart.Type);
            
            // Set the picture
            System.Drawing.Image image1 = (System.Drawing.Image)new Bitmap(dataDir + "aspose-logo.jpg");
            IPPImage imgx1 = presentation.Images.AddImage(image1);

            // Set the picture
            System.Drawing.Image image2 = (System.Drawing.Image)new Bitmap(dataDir + "Tulips.jpg");
            IPPImage imgx2 = presentation.Images.AddImage(image2);

            // Take first chart series
            IChartSeries series = chart.ChartData.Series[0];

            // Add new point (1:3) there.
            IChartDataPoint point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, (double)4.5));
            point.Marker.Format.Fill.FillType = FillType.Picture;
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx1;

            point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, (double)2.5));
            point.Marker.Format.Fill.FillType = FillType.Picture;
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx2;

            point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, (double)3.5));
            point.Marker.Format.Fill.FillType = FillType.Picture;
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx1;

            point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 4, 1, (double)4.5));
            point.Marker.Format.Fill.FillType = FillType.Picture;
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx2;

            // Changing the chart series marker
            series.Marker.Size = 15;

            // Write presentation to disk
            presentation.Save(dataDir + "MarkOptions_out.pptx", SaveFormat.Pptx);
            //ExEnd:SetMarkerOptions
        }
    }
}