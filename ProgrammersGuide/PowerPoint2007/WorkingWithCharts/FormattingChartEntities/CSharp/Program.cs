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

namespace FormattingChartEntities
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

            //Instantiating presentation
            PresentationEx pres = new PresentationEx();

            //Accessing the first slide
            SlideEx slide = pres.Slides[0];

            //Adding the sample chart
            ChartEx chart = slide.Shapes.AddChart(ChartTypeEx.LineWithMarkers, 50, 50, 500, 400);

            //Setting Chart Titile
            chart.HasTitle = true;
            PortionEx chartTitle = chart.ChartTitle.Text.Paragraphs[0].Portions[0];
            chartTitle.Text = "Sample Chart";
            chartTitle.PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
            chartTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray;
            chartTitle.PortionFormat.FontHeight = 20;
            chartTitle.PortionFormat.FontBold = NullableBool.True;
            chartTitle.PortionFormat.FontItalic = NullableBool.True;

            //Setting Major grid lines format for value axis
            chart.ValueAxis.MajorGridLines.FillFormat.FillType = FillTypeEx.Solid;
            chart.ValueAxis.MajorGridLines.FillFormat.SolidFillColor.Color = Color.Blue;
            chart.ValueAxis.MajorGridLines.Width = 5;
            chart.ValueAxis.MajorGridLines.DashStyle = LineDashStyleEx.DashDot;

            //Setting Minor grid lines format for value axis
            chart.ValueAxis.MinorGridLines.FillFormat.FillType = FillTypeEx.Solid;
            chart.ValueAxis.MinorGridLines.FillFormat.SolidFillColor.Color = Color.Red;
            chart.ValueAxis.MinorGridLines.Width = 3;

            //Setting value axis number format
            chart.ValueAxis.SourceLinked = false;
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            chart.ValueAxis.NumberFormat = "0.0%";

            //Setting chart maximum, minimum values
            chart.ValueAxis.IsAutomaticMajorUnit = false;
            chart.ValueAxis.IsAutomaticMaxValue = false;
            chart.ValueAxis.IsAutomaticMinorUnit = false;
            chart.ValueAxis.IsAutomaticMinValue = false;

            chart.ValueAxis.MaxValue = 15f;
            chart.ValueAxis.MinValue = -2f;
            chart.ValueAxis.MinorUnit = 0.5f;
            chart.ValueAxis.MajorUnit = 2.0f;

            //Setting Value Axis Text Properties
            TextFrameEx txtVal = chart.ValueAxis.TextProperties;
            txtVal.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontBold = NullableBool.True;
            txtVal.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontHeight = 16;
            txtVal.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontItalic = NullableBool.True;
            txtVal.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillTypeEx.Solid; ;
            txtVal.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.DarkGreen;
            txtVal.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.LatinFont = new FontDataEx("Times New Roman");

            //Setting value axis title
            chart.ValueAxis.HasTitle = true;
            PortionEx valtitle = chart.ValueAxis.Title.Text.Paragraphs[0].Portions[0];
            valtitle.Text = "Primary Axis";
            valtitle.PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
            valtitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray;
            valtitle.PortionFormat.FontHeight = 20;
            valtitle.PortionFormat.FontBold = NullableBool.True;
            valtitle.PortionFormat.FontItalic = NullableBool.True;

            //Setting value axis line format
            chart.ValueAxis.Format.Line.Width = 10;
            chart.ValueAxis.Format.Line.FillFormat.FillType = FillTypeEx.Solid;
            chart.ValueAxis.Format.Line.FillFormat.SolidFillColor.Color = Color.Red;

            //Setting Major grid lines format for Category axis
            chart.CategoryAxis.MajorGridLines.FillFormat.FillType = FillTypeEx.Solid;
            chart.CategoryAxis.MajorGridLines.FillFormat.SolidFillColor.Color = Color.Green;
            chart.CategoryAxis.MajorGridLines.Width = 5;

            //Setting Minor grid lines format for Category axis
            chart.CategoryAxis.MinorGridLines.FillFormat.FillType = FillTypeEx.Solid;
            chart.CategoryAxis.MinorGridLines.FillFormat.SolidFillColor.Color = Color.Yellow;
            chart.CategoryAxis.MinorGridLines.Width = 3;

            //Setting Category Axis Text Properties
            TextFrameEx txtCat = chart.CategoryAxis.TextProperties;
            txtCat.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontBold = NullableBool.True;
            txtCat.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontHeight = 16;
            txtCat.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontItalic = NullableBool.True;
            txtCat.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillTypeEx.Solid; ;
            txtCat.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;
            txtCat.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.LatinFont = new FontDataEx("Arial");

            //Setting Category Titile
            chart.CategoryAxis.HasTitle = true;
            PortionEx catTitle = chart.CategoryAxis.Title.Text.Paragraphs[0].Portions[0];
            catTitle.Text = "Sample Category";
            catTitle.PortionFormat.FillFormat.FillType = FillTypeEx.Solid;
            catTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray;
            catTitle.PortionFormat.FontHeight = 20;
            catTitle.PortionFormat.FontBold = NullableBool.True;
            catTitle.PortionFormat.FontItalic = NullableBool.True;

            //Setting category axis lable position
            chart.CategoryAxis.TickLabelPosition = TickLabelPositionType.Low;

            //Setting category axis lable rotation angle
            chart.CategoryAxis.RotationAngle = 45;

            //Setting Legends Text Properties
            TextFrameEx txtleg = chart.Legend.TextProperties;
            txtleg.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontBold = NullableBool.True;
            txtleg.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontHeight = 16;
            txtleg.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontItalic = NullableBool.True;
            txtleg.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillTypeEx.Solid; ;
            txtleg.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.DarkRed;

            //Set show chart legends without overlapping chart

            chart.Legend.Overlay = true;

            //Setting secondary value axis
            chart.SecondValueAxis.IsVisible = true;
            chart.SecondValueAxis.Format.Line.Style = LineStyleEx.ThickBetweenThin;
            chart.SecondValueAxis.Format.Line.Width = 20;

            //Setting secondary value axis Number format
            chart.SecondValueAxis.SourceLinked = false;
            chart.SecondValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
            chart.SecondValueAxis.NumberFormat = "0.0%";

            //Setting chart maximum, minimum values
            chart.SecondValueAxis.IsAutomaticMajorUnit = false;
            chart.SecondValueAxis.IsAutomaticMaxValue = false;
            chart.SecondValueAxis.IsAutomaticMinorUnit = false;
            chart.SecondValueAxis.IsAutomaticMinValue = false;

            chart.SecondValueAxis.MaxValue = 20f;
            chart.SecondValueAxis.MinValue = -5f;
            chart.SecondValueAxis.MinorUnit = 0.5f;
            chart.SecondValueAxis.MajorUnit = 2.0f;

            //Ploting first series on secondary value axis
            chart.ChartData.Series[0].PlotOnSecondAxis = true;

            //Setting chart back wall color
            chart.ChartFormat.Fill.FillType = FillTypeEx.Solid;
            chart.ChartFormat.Fill.SolidFillColor.Color = Color.Orange;

            //Setting Plot area color
            chart.PlotArea.Format.Fill.FillType = FillTypeEx.Solid;
            chart.PlotArea.Format.Fill.SolidFillColor.Color = Color.LightCyan;

            //Save Presentation
            pres.Write(dataDir + "ChartAxis.pptx");

        }
    }
}