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
using Aspose.Slides.Pptx.Charts;
using System.Drawing;

namespace SettingPieChartSectorColors
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

            //Instantiate PresentationEx class that represents PPTX file
            PresentationEx pres = new PresentationEx();

            //Access first slide
            SlideEx sld = pres.Slides[0];

            // Add chart with default data
            Aspose.Slides.Pptx.ChartEx chart = sld.Shapes.AddChart(ChartTypeEx.Pie, 100, 100, 400, 400);

            //Setting chart Title
            chart.ChartTitle.Text.Text = "Sample Title";
            chart.ChartTitle.Text.CenterText = true;
            chart.ChartTitle.Height = 20;
            chart.HasTitle = true;

            //Set first series to Show Values
            chart.ChartData.Series[0].Labels.ShowValue = true;

            //Setting the index of chart data sheet 
            int defaultWorksheetIndex = 0;

            //Getting the chart data worksheet
            ChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

            //Delete default generated series and categories

            chart.ChartData.Series.Clear();
            chart.ChartData.Categories.Clear();

            //Adding new categories
            chart.ChartData.Categories.Add(fact.GetCell(0, 1, 0, "First Qtr"));
            chart.ChartData.Categories.Add(fact.GetCell(0, 2, 0, "2nd Qtr"));
            chart.ChartData.Categories.Add(fact.GetCell(0, 3, 0, "3rd Qtr"));

            //Adding new series
            int Id = chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, "Series 1"), chart.Type);

            //Accessing added series
            ChartSeriesEx series = chart.ChartData.Series[Id];

            //Now populating series data
            series.Values.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, 20));
            series.Values.Add(fact.GetCell(defaultWorksheetIndex, 2, 1, 50));
            series.Values.Add(fact.GetCell(defaultWorksheetIndex, 3, 1, 30));

            //Adding new points and setting sector color
            series.IsColorVaried = true;
            ChartPointEx point = new ChartPointEx(series);
            point.Index = 0;
            point.Format.Fill.FillType = FillTypeEx.Solid;
            point.Format.Fill.SolidFillColor.Color = Color.Cyan;
            //Setting Sector border
            point.Format.Line.FillFormat.FillType = FillTypeEx.Solid;
            point.Format.Line.FillFormat.SolidFillColor.Color = Color.Gray;
            point.Format.Line.Width = 3.0;
            point.Format.Line.Style = LineStyleEx.ThinThick;
            point.Format.Line.DashStyle = LineDashStyleEx.DashDot;



            ChartPointEx point1 = new ChartPointEx(series);
            point1.Index = 1;
            point1.Format.Fill.FillType = FillTypeEx.Solid;
            point1.Format.Fill.SolidFillColor.Color = Color.Brown;

            //Setting Sector border
            point1.Format.Line.FillFormat.FillType = FillTypeEx.Solid;
            point1.Format.Line.FillFormat.SolidFillColor.Color = Color.Blue;
            point1.Format.Line.Width = 3.0;
            point1.Format.Line.Style = LineStyleEx.Single;
            point1.Format.Line.DashStyle = LineDashStyleEx.LargeDashDot;

            ChartPointEx point2 = new ChartPointEx(series);
            point2.Index = 2;
            point2.Format.Fill.FillType = FillTypeEx.Solid;
            point2.Format.Fill.SolidFillColor.Color = Color.Coral;

            //Setting Sector border
            point2.Format.Line.FillFormat.FillType = FillTypeEx.Solid;
            point2.Format.Line.FillFormat.SolidFillColor.Color = Color.Red;
            point2.Format.Line.Width = 2.0;
            point2.Format.Line.Style = LineStyleEx.ThinThin;
            point2.Format.Line.DashStyle = LineDashStyleEx.LargeDashDotDot;

            //Adding Series Points
            series.Points.Add(point);
            series.Points.Add(point1);
            series.Points.Add(point2);

            //Create custom labels for each of categories for new series

            DataLabelEx lbl = new DataLabelEx(series);
            // lbl.ShowCategoryName = true;
            lbl.ShowValue = true;
            lbl.Id = 0;
            series.Labels.Add(lbl);

            //Showing Leader Lines for Chart
            series.Labels.ShowLeaderLines = true;

            //Setting Rotation Angle for Pie Chart Sectors
            chart.ChartData.Series[0].FirstSliceAngle = 180;

            // Save presentation with chart
            pres.Write(dataDir + "AsposeChart.pptx");

        }
    }
}