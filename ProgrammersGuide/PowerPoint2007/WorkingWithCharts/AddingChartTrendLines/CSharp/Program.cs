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

namespace AddingChartTrendLines
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

            //Creating empty presentation
            PresentationEx pres = new PresentationEx();

            //Creating a clustered column chart
            ChartEx chart = pres.Slides[0].Shapes.AddChart(ChartTypeEx.ClusteredColumn, 20, 20, 500, 400);

            //Adding Exponential trend line for chart series 1
            TrendlineEx tredLineExp = new TrendlineEx(chart.ChartData.Series[0]);
            tredLineExp.TrendlineType = TrendlineTypeEx.Exponential;
            tredLineExp.DisplayEquation = false;
            tredLineExp.DisplayRSquaredValue = false;
            chart.ChartData.Series[0].TrendLines.Add(tredLineExp);

            //Adding Linear trend line for chart series 1
            TrendlineEx tredLineLin = new TrendlineEx(chart.ChartData.Series[0]);
            tredLineLin.TrendlineType = TrendlineTypeEx.Linear;
            tredLineLin.Format.Line.FillFormat.FillType = FillTypeEx.Solid;
            tredLineLin.Format.Line.FillFormat.SolidFillColor.Color = Color.Red;
            chart.ChartData.Series[0].TrendLines.Add(tredLineLin);


            //Adding Logarithmic trend line for chart series 2
            TrendlineEx tredLineLog = new TrendlineEx(chart.ChartData.Series[1]);
            tredLineLog.TrendlineType = TrendlineTypeEx.Logarithmic;

            tredLineLog.TextFrame.Text = "New log trend line";
            chart.ChartData.Series[1].TrendLines.Add(tredLineLog);

            //Adding MovingAverage trend line for chart series 2
            TrendlineEx tredLineMovAvg = new TrendlineEx(chart.ChartData.Series[1]);
            tredLineMovAvg.TrendlineType = TrendlineTypeEx.MovingAverage;
            tredLineMovAvg.Period = 3;
            tredLineMovAvg.TrendlineName = "New TrendLine Name";
            chart.ChartData.Series[1].TrendLines.Add(tredLineMovAvg);

            //Adding Polynomial trend line for chart series 3
            TrendlineEx tredLinePol = new TrendlineEx(chart.ChartData.Series[2]);
            tredLinePol.TrendlineType = TrendlineTypeEx.Polynomial;
            tredLinePol.Forward = 1;
            tredLinePol.Order = 3;
            chart.ChartData.Series[2].TrendLines.Add(tredLinePol);

            //Adding Power trend line for chart series 3
            TrendlineEx tredLinePower = new TrendlineEx(chart.ChartData.Series[2]);
            tredLinePower.TrendlineType = TrendlineTypeEx.Power;
            tredLinePower.Backward = 1;
            chart.ChartData.Series[2].TrendLines.Add(tredLinePower);

            //Saving presentation
            pres.Write(dataDir + "TrendLines.pptx");

            
        }
    }
}