using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CSharp.Presentations.Properties
{
	public class SupportForInterrupt
	{

		//ExStart:SupportForInterrupt
		public static void Run()
		{
			string dataDir = RunExamples.GetDataDir_PresentationProperties();
			Action<InterruptionToken> action = (InterruptionToken token) =>
			{
				using (Presentation pres = new Presentation("pres.pptx", new LoadOptions { InterruptionToken = token }))
				{
					pres.Slides[0].GetThumbnail(new Size(960, 720));
					pres.Save("pres.ppt", SaveFormat.Ppt);
				}
			};

			InterruptionTokenSource tokenSource = new InterruptionTokenSource();
			Run(action, tokenSource.Token); // run action in a separate thread from the pool

			Thread.Sleep(5000); // some work

			tokenSource.Interrupt(); // we don't need the result of an interruptable action
		}

		private static void Run(Action<InterruptionToken> action, IInterruptionToken token)
		{
			throw new NotImplementedException();
		}

		static void Run(Action<InterruptionToken> action, InterruptionToken token)
		{
			Task.Run(() => { action(token); });
		}
		//ExEnd:SupportForInterrupt
	}
}
