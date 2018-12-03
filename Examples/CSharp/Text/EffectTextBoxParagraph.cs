using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
	class EffectTextBoxParagraph
	{
		public static void Run()
		{
			//ExStart:EffectTextBoxParagraph
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir + "Test.pptx"))
			{
				ISequence sequence = pres.Slides[0].Timeline.MainSequence;
				IAutoShape autoShape = (IAutoShape)pres.Slides[0].Shapes[1];

				foreach (IParagraph paragraph in autoShape.TextFrame.Paragraphs)
				{
					IEffect[] effects = sequence.GetEffectsByParagraph(paragraph);

					if (effects.Length > 0)
						Console.WriteLine("Paragraph \"" + paragraph.Text + "\" has " + effects[0].Type + " effect.");
				}
			}
			//ExEnd:EffectTextBoxParagraph
		}
		}
}
