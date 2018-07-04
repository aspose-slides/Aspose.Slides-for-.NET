using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class SupportForPrecisionOfData
	{
         	public static void Run()
			{
			//ExStart:SupportForPrecisionOfData
			// The path to the documents directory.
			    string dataDir = RunExamples.GetDataDir_Charts();

			    using (Presentation pres = new Presentation())
			    {
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Line, 50, 50, 450, 300);
				chart.HasDataTable = true;
				chart.ChartData.Series[0].NumberFormatOfValues = "#,##0.00";

				pres.Save(dataDir + "ErrorBarsCustomValues_out.pptx", SaveFormat.Pptx);

     			}
			//ExEnd:SupportForPrecisionOfData
		       }
	        }
		}
	


