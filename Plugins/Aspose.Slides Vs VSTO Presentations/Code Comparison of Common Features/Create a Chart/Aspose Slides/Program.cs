using Aspose.Slides.Pptx;
using Aspose.Slides.Pptx.Charts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create empty presentation
            using (PresentationEx pres = new PresentationEx())
            {

                //Accessing first slide
                SlideEx slide = pres.Slides[0];

                //Addding default chart
                ChartEx ppChart = slide.Shapes.AddChart(ChartTypeEx.ClusteredColumn3D, 20F, 30F, 400F, 300F);

                //Getting Chart data
                ChartDataEx chartData = ppChart.ChartData;

                //Removing Extra default series
                chartData.Series.RemoveAt(1);
                chartData.Series.RemoveAt(1);

                //Modifying chart categories names
                chartData.Categories[0].ChartDataCell.Value = "Bikes";
                chartData.Categories[1].ChartDataCell.Value = "Accessories";
                chartData.Categories[2].ChartDataCell.Value = "Repairs";
                chartData.Categories[3].ChartDataCell.Value = "Clothing";

                //Modifying chart series values for first category
                chartData.Series[0].Values[0].Value = 1000;
                chartData.Series[0].Values[1].Value = 2500;
                chartData.Series[0].Values[2].Value = 4000;
                chartData.Series[0].Values[3].Value = 3000;

                //Setting Chart title
                ppChart.HasTitle = true;
                ppChart.ChartTitle.Text.Text = "2007 Sales";
                PortionFormatEx format = ppChart.ChartTitle.Text.Paragraphs[0].Portions[0].PortionFormat;
                format.FontItalic = NullableBool.True;
                format.FontHeight = 18;
                format.FillFormat.FillType = FillTypeEx.Solid;
                format.FillFormat.SolidFillColor.Color = Color.Black;


                //Setting Axis values
                ppChart.ValueAxis.IsAutomaticMaxValue = false;
                ppChart.ValueAxis.IsAutomaticMinValue = false;
                ppChart.ValueAxis.IsAutomaticMajorUnit = false;
                ppChart.ValueAxis.IsAutomaticMinorUnit = false;

                ppChart.ValueAxis.MaxValue = 4000.0F;
                ppChart.ValueAxis.MinValue = 0.0F;
                ppChart.ValueAxis.MajorUnit = 2000.0F;
                ppChart.ValueAxis.MinorUnit = 1000.0F;
                ppChart.ValueAxis.TickLabelPosition = TickLabelPositionType.NextTo;

                //Setting Chart rotation
                ppChart.Rotation3D.RotationX = 15;
                ppChart.Rotation3D.RotationY = 20;

                //Saving Presentation
                pres.Write("AsposeSampleChart.pptx");
            }
        }
    }
}
