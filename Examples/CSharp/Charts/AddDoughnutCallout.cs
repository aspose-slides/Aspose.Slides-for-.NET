using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;

namespace CSharp.Charts
{
	class AddDoughnutCallout
	{
		public static void Run()
		{
			//ExStart:AddDoughnutCallout
			string dataDir = RunExamples.GetDataDir_Charts();
			Presentation pres = new Presentation(dataDir+"testc.pptx");
			ISlide slide = pres.Slides[0];
			IChart chart = slide.Shapes.AddChart(ChartType.Doughnut, 10, 10, 500, 500, false);
			IChartDataWorkbook workBook = chart.ChartData.ChartDataWorkbook;
			chart.ChartData.Series.Clear();
			chart.ChartData.Categories.Clear();
			chart.HasLegend = false;
			int seriesIndex = 0;
			while (seriesIndex < 15)
			{
				IChartSeries series = chart.ChartData.Series.Add(workBook.GetCell(0, 0, seriesIndex + 1, "SERIES " + seriesIndex), chart.Type);
				series.Explosion = 0;
				series.ParentSeriesGroup.DoughnutHoleSize = (byte)20;
				series.ParentSeriesGroup.FirstSliceAngle = 351;
				seriesIndex++;
			}
			int categoryIndex = 0;
			while (categoryIndex < 15)
			{
				chart.ChartData.Categories.Add(workBook.GetCell(0, categoryIndex + 1, 0, "CATEGORY " + categoryIndex));
				int i = 0;
				while (i < chart.ChartData.Series.Count)
				{
					IChartSeries iCS = chart.ChartData.Series[i];
					IChartDataPoint dataPoint = iCS.DataPoints.AddDataPointForDoughnutSeries(workBook.GetCell(0, categoryIndex + 1, i + 1, 1));
					dataPoint.Format.Fill.FillType = FillType.Solid;
					dataPoint.Format.Line.FillFormat.FillType = FillType.Solid;
					dataPoint.Format.Line.FillFormat.SolidFillColor.Color = Color.White;
					dataPoint.Format.Line.Width = 1;
					dataPoint.Format.Line.Style = LineStyle.Single;
					dataPoint.Format.Line.DashStyle = LineDashStyle.Solid;
					if (i == chart.ChartData.Series.Count - 1)
					{
						IDataLabel lbl = dataPoint.Label;
						lbl.TextFormat.TextBlockFormat.AutofitType = TextAutofitType.Shape;
						lbl.DataLabelFormat.TextFormat.PortionFormat.FontBold = NullableBool.True;
						lbl.DataLabelFormat.TextFormat.PortionFormat.LatinFont = new FontData("DINPro-Bold");
						lbl.DataLabelFormat.TextFormat.PortionFormat.FontHeight = 12;
						lbl.DataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
						lbl.DataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.LightGray;
						lbl.DataLabelFormat.Format.Line.FillFormat.SolidFillColor.Color = Color.White;
						lbl.DataLabelFormat.ShowValue = false;
						lbl.DataLabelFormat.ShowCategoryName = true;
						lbl.DataLabelFormat.ShowSeriesName = false;
						//lbl.DataLabelFormat.ShowLabelAsDataCallout = true;
						lbl.DataLabelFormat.ShowLeaderLines = true;
						lbl.DataLabelFormat.ShowLabelAsDataCallout = false;
						chart.ValidateChartLayout();
						lbl.AsILayoutable.X = (float)lbl.AsILayoutable.X + (float)0.5;
						lbl.AsILayoutable.Y = (float)lbl.AsILayoutable.Y + (float)0.5;
					}
					i++;
				}
				categoryIndex++;
			}
			pres.Save(dataDir+"chart.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

		}
		//ExEnd:AddDoughnutCallout
	}
}