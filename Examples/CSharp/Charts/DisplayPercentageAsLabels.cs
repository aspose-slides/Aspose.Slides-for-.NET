using System;
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
    public class DisplayPercentageAsLabels
    {
        public static void Run()
        {
            //ExStart:DisplayPercentageAsLabels
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            ISlide slide = presentation.Slides[0];
            IChart chart = slide.Shapes.AddChart(ChartType.StackedColumn, 20, 20, 400, 400);
            IChartSeries series = chart.ChartData.Series[0];
            IChartCategory cat;
            double[] total_for_Cat = new double[chart.ChartData.Categories.Count];
            for (int k = 0; k < chart.ChartData.Categories.Count; k++)
            {
                cat = chart.ChartData.Categories[k];

                for (int i = 0; i < chart.ChartData.Series.Count; i++)
                {
                    total_for_Cat[k] = total_for_Cat[k] + Convert.ToDouble(chart.ChartData.Series[i].DataPoints[k].Value.Data);
                }
            }

            double dataPontPercent = 0f;

            for (int x = 0; x < chart.ChartData.Series.Count; x++)
            {
                series = chart.ChartData.Series[x];
                series.Labels.DefaultDataLabelFormat.ShowLegendKey = false;

                for (int j = 0; j < series.DataPoints.Count; j++)
                {
                    IDataLabel lbl = series.DataPoints[j].Label;
                    dataPontPercent = (Convert.ToDouble(series.DataPoints[j].Value.Data) / total_for_Cat[j]) * 100;

                    IPortion port = new Portion();
                    port.Text = String.Format("{0:F2} %", dataPontPercent);
                    port.PortionFormat.FontHeight = 8f;
                    lbl.TextFrameForOverriding.Text = "";
                    IParagraph para = lbl.TextFrameForOverriding.Paragraphs[0];
                    para.Portions.Add(port);

                    lbl.DataLabelFormat.ShowSeriesName = false;
                    lbl.DataLabelFormat.ShowPercentage = false;
                    lbl.DataLabelFormat.ShowLegendKey = false;
                    lbl.DataLabelFormat.ShowCategoryName = false;
                    lbl.DataLabelFormat.ShowBubbleSize = false;

                }

            }

            // Save presentation with chart
            presentation.Save(dataDir + "DisplayPercentageAsLabels_out.pptx", SaveFormat.Pptx);
            //ExEnd:DisplayPercentageAsLabels
        }
    }
}