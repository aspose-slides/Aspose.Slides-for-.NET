using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
Using IChartData.GetRange() method example.
*/
namespace CSharp.Charts
{
    class Chart_GetRange
    {
        public static void Run()
        {
            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 10, 10, 400, 300);
                string result = chart.ChartData.GetRange();
                Console.WriteLine("GetRange result : {0}", result);
            }
        }
    }
}
