using System.IO;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class ChartTrendLines
    {
        public static void Run()
        {
            //ExStart:ChartTrendLines
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Creating empty presentation
            Presentation pres = new Presentation();

            // Creating a clustered column chart
            IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 20, 20, 500, 400);

            // Adding ponential trend line for chart series 1
            ITrendline tredLinep = chart.ChartData.Series[0].TrendLines.Add(TrendlineType.Exponential);
            tredLinep.DisplayEquation = false;
            tredLinep.DisplayRSquaredValue = false;

            // Adding Linear trend line for chart series 1
            ITrendline tredLineLin = chart.ChartData.Series[0].TrendLines.Add(TrendlineType.Linear);
            tredLineLin.TrendlineType = TrendlineType.Linear;
            tredLineLin.Format.Line.FillFormat.FillType = FillType.Solid;
            tredLineLin.Format.Line.FillFormat.SolidFillColor.Color = Color.Red;


            // Adding Logarithmic trend line for chart series 2
            ITrendline tredLineLog = chart.ChartData.Series[1].TrendLines.Add(TrendlineType.Logarithmic);
            tredLineLog.TrendlineType = TrendlineType.Logarithmic;
            tredLineLog.AddTextFrameForOverriding("New log trend line");

            // Adding MovingAverage trend line for chart series 2
            ITrendline tredLineMovAvg = chart.ChartData.Series[1].TrendLines.Add(TrendlineType.MovingAverage);
            tredLineMovAvg.TrendlineType = TrendlineType.MovingAverage;
            tredLineMovAvg.Period = 3;
            tredLineMovAvg.TrendlineName = "New TrendLine Name";

            // Adding Polynomial trend line for chart series 3
            ITrendline tredLinePol = chart.ChartData.Series[2].TrendLines.Add(TrendlineType.Polynomial);
            tredLinePol.TrendlineType = TrendlineType.Polynomial;
            tredLinePol.Forward = 1;
            tredLinePol.Order = 3;

            // Adding Power trend line for chart series 3
            ITrendline tredLinePower = chart.ChartData.Series[1].TrendLines.Add(TrendlineType.Power);
            tredLinePower.TrendlineType = TrendlineType.Power;
            tredLinePower.Backward = 1;

            // Saving presentation
            pres.Save(dataDir + "ChartTrendLines_out.pptx", SaveFormat.Pptx);
            //ExEnd:ChartTrendLines
        }
    }
}