using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class FontPropertiesForInvidualLegend
	{
		public static void Run()
		{

			//ExStart:FontPropertiesForInvidualLegend
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
          {
               IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 600, 400);

				IChartTextFormat tf = chart.Legend.Entries[1].TextFormat;

				tf.PortionFormat.FontBold = NullableBool.True;

				tf.PortionFormat.FontHeight = 20;

				tf.PortionFormat.FontItalic = NullableBool.True;

				tf.PortionFormat.FillFormat.FillType = FillType.Solid; ;

				tf.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;
				pres.Save(dataDir+"output.pptx", SaveFormat.Pptx);

			}
			//ExEnd:FontPropertiesForInvidualLegend
		}
	}
}
