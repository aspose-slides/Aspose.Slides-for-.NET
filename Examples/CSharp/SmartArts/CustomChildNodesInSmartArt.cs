using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.SmartArts
{
	class CustomChildNodesInSmartArt
	{
		public static void Run()
		{
			//ExStart:CustomChildNodesInSmartArt
			string dataDir = RunExamples.GetDataDir_SmartArts();

			// Load the desired the presentation
			Presentation pres = new Presentation(dataDir + "AccessChildNodes.pptx");

			{
				ISmartArt smart = pres.Slides[0].Shapes.AddSmartArt(20, 20, 600, 500, SmartArtLayoutType.OrganizationChart);

				// Move SmartArt shape to new position
				ISmartArtNode node = smart.AllNodes[1];
				ISmartArtShape shape = node.Shapes[1];
				shape.X += (shape.Width * 2);
				shape.Y -= (shape.Height / 2);

				// Change SmartArt shape's widths
				node = smart.AllNodes[2];
				shape = node.Shapes[1];
				shape.Width += (shape.Width / 2);

				// Change SmartArt shape's height
				node = smart.AllNodes[3];
				shape = node.Shapes[1];
				shape.Height += (shape.Height / 2);

				// Change SmartArt shape's rotation
				node = smart.AllNodes[4];
				shape = node.Shapes[1];
				shape.Rotation = 90;

				pres.Save(dataDir + "SmartArt.pptx", SaveFormat.Pptx);
			}
			//ExEnd:CustomChildNodesInSmartArt
		}
	}
}