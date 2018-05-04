using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Rendering_Printing
{
	public class RenderingEmoji
	{
		//ExStart:RenderingEmoji
		public static void Run()
		{
			string dataDir = RunExamples.GetDataDir_Rendering();

			Presentation pres = new Presentation(dataDir+"input.pptx");

			pres.Save(dataDir+"emoji.pdf",Aspose.Slides.Export.SaveFormat.Pdf);
         }

	}
      //ExEnd:RenderingEmoji
}
