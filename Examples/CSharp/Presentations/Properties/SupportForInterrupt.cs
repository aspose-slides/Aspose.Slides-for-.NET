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

            Action<IInterruptionToken> action = (IInterruptionToken token) =>
            {
                LoadOptions options = new LoadOptions { InterruptionToken = token };
                using (Presentation presentation = new Presentation(dataDir + "pres.pptx", options))
                {
                    presentation.Save(dataDir + "pres.ppt", SaveFormat.Ppt);
                }
            };

            InterruptionTokenSource tokenSource = new InterruptionTokenSource();
            Run(action, tokenSource.Token); // run action in a separate thread
            Thread.Sleep(10000);            // timeout
            tokenSource.Interrupt();        // stop conversion


        }
        private static void Run(Action<IInterruptionToken> action, IInterruptionToken token)
        {
            Task.Run(() => { action(token); });
        }

        //ExEnd:SupportForInterrupt

    }


}
