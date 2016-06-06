//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
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

namespace CSharp.Charts
{
    public class ChartTrendLines
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Creating empty presentation
            Presentation pres = new Presentation();

            //Creating a clustered column chart
            IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 20, 20, 500, 400);

            //Adding ponential trend line for chart series 1
            ITrendline tredLinep = chart.ChartData.Series[0].TrendLines.Add(TrendlineType.Exponential);
            tredLinep.DisplayEquation = false;
            tredLinep.DisplayRSquaredValue = false;

            //Adding Linear trend line for chart series 1
            ITrendline tredLineLin = chart.ChartData.Series[0].TrendLines.Add(TrendlineType.Linear);
            tredLineLin.TrendlineType = TrendlineType.Linear;
            tredLineLin.Format.Line.FillFormat.FillType = FillType.Solid;
            tredLineLin.Format.Line.FillFormat.SolidFillColor.Color = Color.Red;


            //Adding Logarithmic trend line for chart series 2
            ITrendline tredLineLog = chart.ChartData.Series[1].TrendLines.Add(TrendlineType.Logarithmic);
            tredLineLog.TrendlineType = TrendlineType.Logarithmic;
            tredLineLog.AddTextFrameForOverriding("New log trend line");

            //Adding MovingAverage trend line for chart series 2
            ITrendline tredLineMovAvg = chart.ChartData.Series[1].TrendLines.Add(TrendlineType.MovingAverage);
            tredLineMovAvg.TrendlineType = TrendlineType.MovingAverage;
            tredLineMovAvg.Period = 3;
            tredLineMovAvg.TrendlineName = "New TrendLine Name";

            //Adding Polynomial trend line for chart series 3
            ITrendline tredLinePol = chart.ChartData.Series[2].TrendLines.Add(TrendlineType.Polynomial);
            tredLinePol.TrendlineType = TrendlineType.Polynomial;
            tredLinePol.Forward = 1;
            tredLinePol.Order = 3;

            //Adding Power trend line for chart series 3
            ITrendline tredLinePower = chart.ChartData.Series[1].TrendLines.Add(TrendlineType.Power);
            tredLinePower.TrendlineType = TrendlineType.Power;
            tredLinePower.Backward = 1;

            //Saving presentation
            pres.Save(dataDir + "ChartTrendLines.pptx", SaveFormat.Pptx);
 
        }
    }
}