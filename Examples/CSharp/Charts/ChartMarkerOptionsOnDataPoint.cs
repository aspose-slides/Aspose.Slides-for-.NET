using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class ChartMarkerOptionsOnDataPoint
	{
		public static void Run()
		{
			//ExStart:ChartMarkerOptionsOnDataPoint
			string dataDir = RunExamples.GetDataDir_Charts();
			Presentation pres = new Presentation(dataDir+"Test.pptx");

			ISlide slide = pres.Slides[0];

			//Creating the default chart
			IChart chart = slide.Shapes.AddChart(ChartType.LineWithMarkers, 0, 0, 400, 400);

			//Getting the default chart data worksheet index
			int defaultWorksheetIndex = 0;

			//Getting the chart data worksheet
			IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

			//Delete demo series
			chart.ChartData.Series.Clear();

			//Add new series
			chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, "Series 1"), chart.Type);

			
			//Set the picture
		    IImage img = Images.FromFile(dataDir + "aspose-logo.jpg");
			IPPImage imgx1 = pres.Images.AddImage(img);

            //Set the picture
		    IImage img2 = Images.FromFile(dataDir + "Tulips.jpg");
			IPPImage imgx2 = pres.Images.AddImage(img2);

			//Take first chart series
			IChartSeries series = chart.ChartData.Series[0];

			//Add new point (1:3) there.
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


			//Changing the chart series marker
			series.Marker.Size = 15;

			pres.Save(dataDir+"AsposeScatterChart.pptx", SaveFormat.Pptx);
		}

		//ExEnd:ChartMarkerOptionsOnDataPoint
	}
  }
 