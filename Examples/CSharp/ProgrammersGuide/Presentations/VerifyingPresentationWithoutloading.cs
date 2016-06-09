using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Slides;

namespace CSharp.ProgrammersGuide.Presentations
{
    class VerifyingPresentationWithoutloading
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            LoadFormat format = PresentationFactory.Instance.GetPresentationInfo(dataDir+"HelloWorld.pptx").LoadFormat;
            // It will return "LoadFormat.Unknown" if the file is other than presentation formats           
        }
    }
}
