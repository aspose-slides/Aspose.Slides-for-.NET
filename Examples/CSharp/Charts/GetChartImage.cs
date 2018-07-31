using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class GetChartImage
	{
		public static void Run()
		{
			//ExStart:GetChartImage
			// The path to the documents directory.
			string dataDir = RunExamples.GetDataDir_Charts();

			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
            {
            	IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 600, 400);
             	Image img = chart.GetThumbnail();
             	img.Save(dataDir+"image.png", ImageFormat.Png);
			}
			//ExEnd:GetChartImage
		}
	}
}
