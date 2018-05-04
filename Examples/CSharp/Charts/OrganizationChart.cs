using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	public class OrganizationChart
	{
		//ExStart:OrganizationChart
		public static void Run()
			 {
			
			// The path to the documents directory.
			string dataDir = RunExamples.GetDataDir_Charts();
				using (Presentation pres = new Presentation(dataDir+"test.pptx"))
				{
					ISmartArt smartArt = pres.Slides[0].Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.PictureOrganizationChart);

					pres.Save(dataDir+"OrganizationChart.pptx", SaveFormat.Pptx);
				}			

			}
		//ExEnd:OrganizationChart
	    }
	}
