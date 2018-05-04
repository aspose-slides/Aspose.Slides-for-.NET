using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Conversion
{
	class ExportShapeToSVG
	{
		//ExStart:ExportShapeToSVG
		public static void Run()
		{
			
			string outSvgFileName = "SingleShape.svg";
			string dataDir = RunExamples.GetDataDir_Conversion();
			using (Presentation pres = new Presentation(dataDir+ "TestExportShapeToSvg.pptx"))
			{
				using (Stream stream = new FileStream(outSvgFileName, FileMode.Create, FileAccess.Write))
				{
					pres.Slides[0].Shapes[0].WriteAsSvg(stream);

					
				}
			
				
			}


		}


		//ExEnd:ExportShapeToSVG
	}
}